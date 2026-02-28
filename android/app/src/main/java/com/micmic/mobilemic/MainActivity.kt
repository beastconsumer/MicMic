package com.micmic.mobilemic

import android.Manifest
import android.content.Intent
import android.content.pm.PackageManager
import android.os.Bundle
import android.widget.TextView
import androidx.appcompat.app.AppCompatActivity
import androidx.core.app.ActivityCompat
import androidx.core.content.ContextCompat
import com.google.android.material.button.MaterialButton

class MainActivity : AppCompatActivity() {

    private lateinit var statusText: TextView
    private lateinit var startButton: MaterialButton
    private lateinit var stopButton: MaterialButton

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        statusText = findViewById(R.id.txtStatus)
        startButton = findViewById(R.id.btnStart)
        stopButton = findViewById(R.id.btnStop)

        startButton.setOnClickListener { ensurePermissionAndStart() }
        stopButton.setOnClickListener { stopCapture() }

        handleExternalCommand(intent)
    }

    override fun onNewIntent(intent: Intent) {
        super.onNewIntent(intent)
        handleExternalCommand(intent)
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
            statusText.text = getString(R.string.status_permission_needed)
        }
    }

    private fun startCapture() {
        val serviceIntent = Intent(this, MicStreamService::class.java).apply {
            action = MicStreamService.ACTION_START
        }
        ContextCompat.startForegroundService(this, serviceIntent)
        statusText.text = getString(R.string.status_recording)
    }

    private fun stopCapture() {
        val serviceIntent = Intent(this, MicStreamService::class.java).apply {
            action = MicStreamService.ACTION_STOP
        }
        startService(serviceIntent)
        statusText.text = getString(R.string.status_stopped)
    }

    companion object {
        private const val REQUEST_AUDIO_PERMISSION = 1001
        const val EXTRA_COMMAND = "command"
    }
}
