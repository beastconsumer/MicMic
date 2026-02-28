package com.micmic.mobilemic

import android.Manifest
import android.content.BroadcastReceiver
import android.content.Context
import android.content.Intent
import android.content.IntentFilter
import android.content.pm.PackageManager
import android.content.res.ColorStateList
import android.os.Bundle
import android.view.View
import android.widget.TextView
import androidx.appcompat.app.AppCompatActivity
import androidx.core.app.ActivityCompat
import androidx.core.content.ContextCompat
import com.google.android.material.button.MaterialButton

class MainActivity : AppCompatActivity() {

    private lateinit var statusText: TextView
    private lateinit var connectionText: TextView
    private lateinit var connectionDot: View
    private lateinit var startButton: MaterialButton
    private lateinit var stopButton: MaterialButton
    private var receiverRegistered = false

    private val stateReceiver = object : BroadcastReceiver() {
        override fun onReceive(context: Context?, intent: Intent?) {
            if (intent?.action != MicStreamService.ACTION_STATE_CHANGED) {
                return
            }
            val state = intent.getStringExtra(MicStreamService.EXTRA_STATE) ?: return
            renderConnectionState(state)
        }
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        statusText = findViewById(R.id.txtStatus)
        connectionText = findViewById(R.id.txtConnectionState)
        connectionDot = findViewById(R.id.viewConnectionDot)
        startButton = findViewById(R.id.btnStart)
        stopButton = findViewById(R.id.btnStop)

        startButton.setOnClickListener { ensurePermissionAndStart() }
        stopButton.setOnClickListener { stopCapture() }

        renderConnectionState(MicStreamService.currentState())
        handleExternalCommand(intent)
    }

    override fun onNewIntent(intent: Intent) {
        super.onNewIntent(intent)
        handleExternalCommand(intent)
    }

    override fun onStart() {
        super.onStart()
        registerStateReceiver()
        renderConnectionState(MicStreamService.currentState())
    }

    override fun onStop() {
        if (receiverRegistered) {
            unregisterReceiver(stateReceiver)
            receiverRegistered = false
        }
        super.onStop()
    }

    private fun handleExternalCommand(intent: Intent?) {
        val command = intent?.getStringExtra(EXTRA_COMMAND)?.lowercase() ?: return
        when (command) {
            "start" -> ensurePermissionAndStart()
            "stop" -> stopCapture()
        }
    }

    private fun ensurePermissionAndStart() {
        val granted = ContextCompat.checkSelfPermission(
            this,
            Manifest.permission.RECORD_AUDIO,
        ) == PackageManager.PERMISSION_GRANTED

        if (granted) {
            startCapture()
            return
        }

        ActivityCompat.requestPermissions(
            this,
            arrayOf(Manifest.permission.RECORD_AUDIO),
            REQUEST_AUDIO_PERMISSION,
        )
    }

    override fun onRequestPermissionsResult(
        requestCode: Int,
        permissions: Array<out String>,
        grantResults: IntArray,
    ) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults)
        if (requestCode != REQUEST_AUDIO_PERMISSION) {
            return
        }

        if (grantResults.isNotEmpty() && grantResults[0] == PackageManager.PERMISSION_GRANTED) {
            startCapture()
        } else {
            renderConnectionState(MicStreamService.STATE_PERMISSION_DENIED)
        }
    }

    private fun startCapture() {
        val serviceIntent = Intent(this, MicStreamService::class.java).apply {
            action = MicStreamService.ACTION_START
        }
        ContextCompat.startForegroundService(this, serviceIntent)
        renderConnectionState(MicStreamService.STATE_CONNECTING)
    }

    private fun stopCapture() {
        val serviceIntent = Intent(this, MicStreamService::class.java).apply {
            action = MicStreamService.ACTION_STOP
        }
        startService(serviceIntent)
        renderConnectionState(MicStreamService.STATE_STOPPED)
    }

    private fun registerStateReceiver() {
        if (receiverRegistered) {
            return
        }
        val filter = IntentFilter(MicStreamService.ACTION_STATE_CHANGED)
        ContextCompat.registerReceiver(
            this,
            stateReceiver,
            filter,
            ContextCompat.RECEIVER_NOT_EXPORTED,
        )
        receiverRegistered = true
    }

    private fun renderConnectionState(state: String) {
        val (dotColorRes, connectionLabelRes, statusLabelRes, streamingLike) = when (state) {
            MicStreamService.STATE_CONNECTED -> StateUi(
                R.color.status_connected,
                R.string.connection_connected,
                R.string.status_recording,
                true,
            )
            MicStreamService.STATE_CONNECTING -> StateUi(
                R.color.status_connecting,
                R.string.connection_connecting,
                R.string.status_connecting,
                true,
            )
            MicStreamService.STATE_RECONNECTING -> StateUi(
                R.color.status_connecting,
                R.string.connection_reconnecting,
                R.string.status_reconnecting,
                true,
            )
            MicStreamService.STATE_STOPPED -> StateUi(
                R.color.status_disconnected,
                R.string.connection_disconnected,
                R.string.status_stopped,
                false,
            )
            MicStreamService.STATE_PERMISSION_DENIED -> StateUi(
                R.color.status_disconnected,
                R.string.connection_disconnected,
                R.string.status_permission_needed,
                false,
            )
            MicStreamService.STATE_ERROR -> StateUi(
                R.color.status_disconnected,
                R.string.connection_disconnected,
                R.string.status_stopped,
                false,
            )
            else -> StateUi(
                R.color.status_disconnected,
                R.string.connection_disconnected,
                R.string.status_ready,
                false,
            )
        }

        val dotColor = ContextCompat.getColor(this, dotColorRes)
        connectionDot.backgroundTintList = ColorStateList.valueOf(dotColor)
        connectionText.text = getString(connectionLabelRes)
        statusText.text = getString(statusLabelRes)
        startButton.isEnabled = !streamingLike
        stopButton.isEnabled = streamingLike
    }

    companion object {
        private const val REQUEST_AUDIO_PERMISSION = 1001
        const val EXTRA_COMMAND = "command"
    }

    private data class StateUi(
        val dotColorRes: Int,
        val connectionLabelRes: Int,
        val statusLabelRes: Int,
        val streamingLike: Boolean,
    )
}
