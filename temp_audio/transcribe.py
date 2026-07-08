import whisper
import os
import sys
import time

# Set ffmpeg path
os.environ["PATH"] += os.pathsep + r"E:\01.NDRS_WORKON\OpenWork\temp_audio\bin"
os.environ["PATH"] += os.pathsep + r"C:\Users\soraw\AppData\Local\Temp\ffmpeg_extract\ffmpeg-8.1.2-essentials_build\bin"

# Copy cached model to whisper's expected location
cache_dir = os.path.expanduser("~/.cache/whisper")
os.makedirs(cache_dir, exist_ok=True)

# Check if medium model already cached by Buzz
buzz_model = os.path.expanduser(r"~\AppData\Local\Buzz\Buzz\Cache\models\whisper\medium.pt")
whisper_model = os.path.join(cache_dir, "medium.pt")

if os.path.exists(buzz_model) and not os.path.exists(whisper_model):
    print(f"Copying Buzz cached model to Whisper cache...")
    import shutil
    shutil.copy2(buzz_model, whisper_model)
    print(f"Copied {os.path.getsize(whisper_model)} bytes")

audio_file = sys.argv[1]
output_dir = sys.argv[2] if len(sys.argv) > 2 else os.path.dirname(audio_file)

print(f"Loading model 'medium'...")
t0 = time.time()
model = whisper.load_model("medium")
print(f"Model loaded in {time.time()-t0:.1f}s")

print(f"Transcribing {audio_file}...")
t0 = time.time()
result = model.transcribe(audio_file, language="en", verbose=False)
print(f"Transcription done in {time.time()-t0:.1f}s")

# Save TXT
txt_path = os.path.join(output_dir, "session1_transcript.txt")
with open(txt_path, "w", encoding="utf-8") as f:
    f.write(result["text"])
print(f"Saved: {txt_path}")

# Save with timestamps
segments_path = os.path.join(output_dir, "session1_segments.txt")
with open(segments_path, "w", encoding="utf-8") as f:
    for seg in result["segments"]:
        start = seg["start"]
        end = seg["end"]
        text = seg["text"].strip()
        f.write(f"[{start:.2f}s -> {end:.2f}s] {text}\n")
print(f"Saved: {segments_path}")

print("DONE")
