import obspython as obs
import os
import subprocess
import time

def script_description():
    return "Sドライブ内を再帰的に検索してクリップボードへ送り、消去します。"

def take_screenshot_and_process():
    print("--- スクショ処理開始 ---")
    obs.obs_frontend_take_screenshot()
    
    # 書き出し完了待ち（念のため1.2秒）
    time.sleep(1.2)
    
    search_dir = "S:\\"
    
    # 修正ポイント: 
    # 1. -Recurse を追加してサブフォルダも探す
    # 2. -Include ではなく -Filter を使うか、拡張子チェックを柔軟にする
    ps_cmd = (
        f"$dir = '{search_dir}'; "
        "$file = Get-ChildItem -Path $dir -File -Recurse | Where-Object {{ $_.Extension -match 'png|jpg|jpeg' }} | Sort-Object LastWriteTime -Descending | Select-Object -First 1; "
        "if ($file) { "
        "  write-host ('対象ファイル発見: ' + $file.FullName); "
        "  try { "
        "    Add-Type -AssemblyName System.Windows.Forms, System.Drawing; "
        "    $img = [System.Drawing.Image]::FromFile($file.FullName); "
        "    [System.Windows.Forms.Clipboard]::SetImage($img); "
        "    $img.Dispose(); "
        "    Remove-Item -LiteralPath $file.FullName -Force; "
        "    write-host '成功: クリップボード転送と削除が完了しました。'; "
        "  } catch { "
        "    write-error ('エラー発生: ' + $_.Exception.Message); "
        "  } "
        "} else { "
        "  write-host 'エラー: Sドライブ内に画像が見つかりませんでした。現在あるファイル:'; "
        "  Get-ChildItem -Path $dir -Recurse | Select-Object Name; "
        "}"
        "[System.Media.SystemSounds]::Asterisk.Play();"
    )
    
    result = subprocess.run(
        ["powershell", "-Command", ps_cmd], 
        capture_output=True,
        text=True,
        creationflags=0x08000000
    )
    
    if result.stdout: print(f"PS出力: {result.stdout.strip()}")
    if result.stderr: print(f"PSエラー: {result.stderr.strip()}")
    print("--- 処理終了 ---")

def script_load(settings):
    global hotkey_id
    hotkey_id = obs.obs_hotkey_register_frontend(
        "obs_to_clip_hotkey", "Screenshot to Clipboard (Fix)", 
        lambda pressed: take_screenshot_and_process() if pressed else None
    )
    save_array = obs.obs_data_get_array(settings, "obs_to_clip_hotkey")
    obs.obs_hotkey_load(hotkey_id, save_array)
    obs.obs_data_array_release(save_array)

def script_save(settings):
    save_array = obs.obs_hotkey_save(hotkey_id)
    obs.obs_data_set_array(settings, "obs_to_clip_hotkey", save_array)
    obs.obs_data_array_release(save_array)