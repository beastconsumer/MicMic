# MICMIC - Celular como microfone no Discord

Projeto com duas partes:

- `android/`: app Android que envia audio do microfone por USB (ADB reverse).
- `desktop/`: app Windows dark, simples, com apenas `Start/Stop` e tudo automatico.

## Requisitos

- Windows com Python 3.10+
- `adb` no PATH (ou `%LOCALAPPDATA%\Android\Sdk\platform-tools\adb.exe`)
- Celular Android com depuracao USB ativa
- Driver de mic virtual instalado no Windows:
  - `Virtual Mic for AudioRelay` + `Virtual Speakers for AudioRelay` (preferido)
  - ou `VB-CABLE` (`CABLE Input` / `CABLE Output`)

## Build APK Android

```powershell
cd android
.\gradlew.bat assembleDebug
```

APK:
- `android\app\build\outputs\apk\debug\app-debug.apk`

## Rodar no Windows

```powershell
cd desktop
python -m pip install -r requirements.txt
python mic_bridge_app.py
```

## Uso (sem configuracao manual de entrada/saida)

1. Conecte o celular via USB.
2. Abra o app no PC e clique `START`.
3. O app:
   - detecta celular;
   - instala APK se faltar;
   - detecta mic virtual automaticamente;
   - define microfone padrao do Windows;
   - inicia stream.
4. No Discord, selecione `MICMIC` se aparecer. Se nao aparecer, selecione `Default`.

## Nome exato `MICMIC` no Discord

- O app tenta renomear o endpoint virtual para `MICMIC`.
- Se o Windows bloquear por permissao, execute o app como administrador.
- Sem driver virtual, o Discord nao consegue enxergar um microfone novo.
