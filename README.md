# MICMIC Studio - Celular como microfone no PC

Projeto com duas partes:

- `android/`: app Android que captura o microfone e envia PCM via socket.
- `desktop/`: painel dark com UX simples para conectar celular, iniciar/parar stream e definir mic padrao no Windows.

## Requisitos

- Windows com Python 3.10+
- `adb` no PATH (ou em `%LOCALAPPDATA%\Android\Sdk\platform-tools\adb.exe`)
- Celular Android com depuracao USB ativa
- Driver de microfone virtual no Windows (uma destas opcoes):
  - `Virtual Mic for AudioRelay` (se ja existir no PC)
  - `VB-CABLE` (`CABLE Input` + `CABLE Output`)

## 1) Build do APK Android

```powershell
cd android
.\gradlew.bat assembleDebug
```

APK gerado em:

- `android\app\build\outputs\apk\debug\app-debug.apk`

## 2) Executar o app desktop (UI dark)

```powershell
cd desktop
python -m pip install -r requirements.txt
python mic_bridge_app.py
```

Ou execute `desktop\run_mic_bridge.bat`.

## 3) Fluxo recomendado (rapido)

1. Conecte o celular por USB.
2. No app desktop, clique `Atualizar diagnostico`.
3. Clique `Instalar/Reinstalar APK` (se necessario).
4. Escolha:
   - `Saida para stream (render)`
   - `Microfone no Windows/Discord (capture)`
5. Mantenha ligado `Ao iniciar, definir microfone padrao automaticamente`.
6. Clique `START STREAM`.

Resultado:

- O app envia comando Start para o celular.
- O Windows recebe audio no microfone virtual escolhido.
- O microfone e definido como padrao (console/multimedia/comunicacoes), entao o Discord pega automaticamente o mic certo.

## 4) Parar

- Clique `STOP` no app desktop.

## Solucao de problemas

- `Celular nao autorizado`: aceite a chave RSA no celular.
- `Nenhum celular conectado`: troque cabo USB e confirme `adb devices`.
- Sem mic virtual na lista: instale VB-CABLE e clique `Atualizar diagnostico`.
- Sem audio: confirme permissao de microfone no app Android.
