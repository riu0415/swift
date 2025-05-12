import win32com.client

editor = win32com.client.Dispatch("AI.Talk.Editor")
editor.Text = "こんにちは、PythonからA.I.VOICEを操作しています。"
editor.Play()  # 再生
# editor.SaveAudioToFile("C:\\output.wav")  # 音声を保存
