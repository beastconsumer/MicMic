package com.micmic.mobilemic

import android.Manifest
import android.app.NotificationChannel
import android.app.NotificationManager
import android.app.Service
import android.content.Context
import android.content.Intent
import android.content.pm.PackageManager
import android.media.AudioFormat
import android.media.AudioRecord
import android.media.MediaRecorder
import android.os.Build
import android.os.IBinder
import androidx.core.app.ActivityCompat
import androidx.core.app.NotificationCompat
import java.io.BufferedOutputStream
import java.io.IOException
import java.net.Socket
import java.util.concurrent.atomic.AtomicBoolean

class MicStreamService : Service() {

    private val running = AtomicBoolean(false)
    private var workerThread: Thread? = null
    private var audioRecord: AudioRecord? = null

    override fun onBind(intent: Intent?): IBinder? = null

    override fun onStartCommand(intent: Intent?, flags: Int, startId: Int): Int {
        when (intent?.action) {
            ACTION_START -> startStreaming()
            ACTION_STOP -> stopStreaming()
        }
        return START_STICKY
    }

    override fun onDestroy() {
        stopStreaming()
        super.onDestroy()
    }

    private fun startStreaming() {
        if (running.get()) {
            publishState(STATE_CONNECTED)
            return
        }

        createNotificationChannel()
        startForeground(NOTIFICATION_ID, buildNotification())

        running.set(true)
        publishState(STATE_CONNECTING)
        workerThread = Thread(::captureLoop, "mic-stream-thread").also { it.start() }
    }

    private fun stopStreaming() {
        running.set(false)
        workerThread?.interrupt()
        workerThread = null

        releaseAudioRecord()

        stopForeground(STOP_FOREGROUND_REMOVE)
        publishState(STATE_STOPPED)
        stopSelf()
    }

    private fun captureLoop() {
        if (ActivityCompat.checkSelfPermission(this, Manifest.permission.RECORD_AUDIO) != PackageManager.PERMISSION_GRANTED) {
            publishState(STATE_PERMISSION_DENIED)
            stopSelf()
            return
        }

        val minBuffer = AudioRecord.getMinBufferSize(
            SAMPLE_RATE,
            AudioFormat.CHANNEL_IN_MONO,
            AudioFormat.ENCODING_PCM_16BIT,
        )
        if (minBuffer <= 0) {
            publishState(STATE_ERROR)
            stopSelf()
            return
        }

        val bufferSize = minBuffer * 2
        val readBuffer = ByteArray(bufferSize)

        audioRecord = AudioRecord(
            MediaRecorder.AudioSource.MIC,
            SAMPLE_RATE,
            AudioFormat.CHANNEL_IN_MONO,
            AudioFormat.ENCODING_PCM_16BIT,
            bufferSize,
        )

        try {
            audioRecord?.startRecording()
        } catch (_: IllegalStateException) {
            publishState(STATE_ERROR)
            releaseAudioRecord()
            stopSelf()
            return
        }

        while (running.get() && !Thread.currentThread().isInterrupted) {
            try {
                publishState(STATE_CONNECTING)
                Socket(HOST, PORT).use { socket ->
                    socket.tcpNoDelay = true
                    publishState(STATE_CONNECTED)
                    BufferedOutputStream(socket.getOutputStream()).use { output ->
                        while (running.get() && !Thread.currentThread().isInterrupted) {
                            val read = audioRecord?.read(readBuffer, 0, readBuffer.size) ?: -1
                            if (read > 0) {
                                output.write(readBuffer, 0, read)
                            } else if (read < 0) {
                                throw IOException("Erro no AudioRecord: $read")
                            }
                        }
                        output.flush()
                    }
                }
            } catch (_: IOException) {
                if (running.get()) {
                    publishState(STATE_RECONNECTING)
                    try {
                        Thread.sleep(RETRY_DELAY_MS)
                    } catch (_: InterruptedException) {
                        Thread.currentThread().interrupt()
                    }
                }
            }
        }

        releaseAudioRecord()
        stopForeground(STOP_FOREGROUND_REMOVE)
        publishState(STATE_STOPPED)
        stopSelf()
    }

    private fun publishState(state: String) {
        if (serviceState == state) {
            return
        }
        serviceState = state
        val intent = Intent(ACTION_STATE_CHANGED).apply {
            setPackage(packageName)
            putExtra(EXTRA_STATE, state)
        }
        sendBroadcast(intent)
    }

    private fun releaseAudioRecord() {
        try {
            audioRecord?.stop()
        } catch (_: IllegalStateException) {
            // Ignore invalid state during teardown.
        }
        audioRecord?.release()
        audioRecord = null
    }

    private fun createNotificationChannel() {
        if (Build.VERSION.SDK_INT < Build.VERSION_CODES.O) {
            return
        }
        val manager = getSystemService(Context.NOTIFICATION_SERVICE) as NotificationManager
        val channel = NotificationChannel(
            CHANNEL_ID,
            getString(R.string.app_name),
            NotificationManager.IMPORTANCE_LOW,
        )
        manager.createNotificationChannel(channel)
    }

    private fun buildNotification() = NotificationCompat.Builder(this, CHANNEL_ID)
        .setSmallIcon(android.R.drawable.ic_btn_speak_now)
        .setContentTitle(getString(R.string.notification_title))
        .setContentText(getString(R.string.notification_text))
        .setOngoing(true)
        .build()

    companion object {
        const val ACTION_START = "com.micmic.mobilemic.action.START"
        const val ACTION_STOP = "com.micmic.mobilemic.action.STOP"
        const val ACTION_STATE_CHANGED = "com.micmic.mobilemic.action.STATE_CHANGED"
        const val EXTRA_STATE = "state"

        const val STATE_STOPPED = "stopped"
        const val STATE_CONNECTING = "connecting"
        const val STATE_CONNECTED = "connected"
        const val STATE_RECONNECTING = "reconnecting"
        const val STATE_PERMISSION_DENIED = "permission_denied"
        const val STATE_ERROR = "error"

        @Volatile
        private var serviceState: String = STATE_STOPPED

        fun currentState(): String = serviceState

        private const val CHANNEL_ID = "mic_stream_channel"
        private const val NOTIFICATION_ID = 2001
        private const val HOST = "127.0.0.1"
        private const val PORT = 28282
        private const val SAMPLE_RATE = 48000
        private const val RETRY_DELAY_MS = 1000L
    }
}
