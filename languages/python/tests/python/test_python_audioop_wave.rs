// Python wave module — WAV file read/write, frame rate, channels
// audioop is deprecated/removed in Python 3.13+, tests focus on wave module
use super::helpers::run_python;

#[test]
fn test_wave_write_and_read_params() {
    let script = r#"
import wave, struct, tempfile, os

path = tempfile.mktemp(suffix='.wav')

# Write a silent mono WAV file (100 frames at 44100 Hz, 16-bit)
with wave.open(path, 'wb') as wf:
    wf.setnchannels(1)
    wf.setsampwidth(2)
    wf.setframerate(44100)
    silence = struct.pack('<' + 'h' * 100, *([0] * 100))
    wf.writeframes(silence)

with wave.open(path, 'rb') as wf:
    print(wf.getnchannels())
    print(wf.getsampwidth())
    print(wf.getframerate())
    print(wf.getnframes())

os.unlink(path)
"#;
    assert_eq!(run_python(script), vec!["1", "2", "44100", "100"]);
}

#[test]
fn test_wave_readframes() {
    let script = r#"
import wave, struct, tempfile, os

path = tempfile.mktemp(suffix='.wav')

values = [100, -100, 200, -200]
with wave.open(path, 'wb') as wf:
    wf.setnchannels(1)
    wf.setsampwidth(2)
    wf.setframerate(8000)
    data = struct.pack('<' + 'h' * len(values), *values)
    wf.writeframes(data)

with wave.open(path, 'rb') as wf:
    frames = wf.readframes(4)
    decoded = struct.unpack('<' + 'h' * 4, frames)
    print(list(decoded))

os.unlink(path)
"#;
    assert_eq!(run_python(script), vec!["[100, -100, 200, -200]"]);
}

#[test]
fn test_wave_stereo_channels() {
    let script = r#"
import wave, struct, tempfile, os

path = tempfile.mktemp(suffix='.wav')

with wave.open(path, 'wb') as wf:
    wf.setnchannels(2)
    wf.setsampwidth(2)
    wf.setframerate(22050)
    # 10 stereo frames
    data = struct.pack('<' + 'hh' * 10, *([0, 0] * 10))
    wf.writeframes(data)

with wave.open(path, 'rb') as wf:
    print(wf.getnchannels())
    print(wf.getnframes())

os.unlink(path)
"#;
    assert_eq!(run_python(script), vec!["2", "10"]);
}

#[test]
fn test_wave_getparams() {
    let script = r#"
import wave, struct, tempfile, os

path = tempfile.mktemp(suffix='.wav')

with wave.open(path, 'wb') as wf:
    wf.setnchannels(1)
    wf.setsampwidth(2)
    wf.setframerate(16000)
    wf.writeframes(b'\x00' * 32)  # 8 silent frames (2 bytes each)

with wave.open(path, 'rb') as wf:
    params = wf.getparams()
    print(params.nchannels)
    print(params.framerate)

os.unlink(path)
"#;
    assert_eq!(run_python(script), vec!["1", "16000"]);
}

#[test]
fn test_wave_comptype_none() {
    let script = r#"
import wave, struct, tempfile, os

path = tempfile.mktemp(suffix='.wav')

with wave.open(path, 'wb') as wf:
    wf.setnchannels(1)
    wf.setsampwidth(2)
    wf.setframerate(8000)
    wf.writeframes(b'\x00' * 2)

with wave.open(path, 'rb') as wf:
    print(wf.getcomptype())
    print(wf.getcompname())

os.unlink(path)
"#;
    assert_eq!(run_python(script), vec!["NONE", "not compressed"]);
}
