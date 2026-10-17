# Direct Rust Canvas Bridge Benchmark

This example depends on the browser transport only: no Vybe compiler, C
adapter, or WASM VM. It uses the same loopback bridge and standard browser
Canvas operations as the SDL benchmark.

```sh
cargo run -p osbrowser --example canvas_bridge_bench -- --frames 120
cargo run -p osbrowser --example canvas_bridge_bench -- --frames 120 --queued
cargo run -p osbrowser --example canvas_bridge_bench -- --frames 120 --copy-only
cargo run -p osbrowser --example canvas_bridge_bench -- --browser safari --frames 16
cargo run -p osbrowser --example canvas_bridge_bench -- --browser firefox --frames 16
cargo run -p osbrowser --example canvas_bridge_bench -- --browser chrome --binary --frames 120
cargo run -p osbrowser --example canvas_bridge_bench -- --browser chrome --binary --queued --frames 120
```

On macOS, `--browser chrome|safari|firefox` selects the browser explicitly.
Default mode prefers Chrome on macOS and otherwise opens the system browser.
Use `--manual` to open the printed URL in a chosen browser, or `--headless` for an owned temporary Chrome instance
that exits after the test. Headless results are not window presentation rates.

Eight cached RGBA checkerboard images animate at 320x200. Each frame uploads
one image, scales it into a 960x600 backbuffer, then copies the backbuffer to
the visible canvas. This matches the SDL texture/backbuffer/window topology.
Image generation, document construction, and the final pixel check are outside
the measured loop. JSON argument cloning is reported separately as preparation;
wire serialization remains part of the bridge cost.

Synchronous mode reports upload and two native `drawImage` round trips.
Queued mode submits all three operations together, then waits on a standard
document-title query as a barrier. Its upload/copy numbers are enqueue costs;
**the barrier and total elapsed time include completion**, so those numbers
must not be mistaken for completed browser work. The initial title-query test
isolates small round-trip overhead. A final `getImageData` verifies that pixels
reached the destination.

Reported FPS measures completed submissions, not display refresh or browser
paint FPS. Compare debug builds with the debug SDL benchmark. Running browsers
or compilers concurrently can change the result.

`--copy-only` uploads all eight images before timing and then animates by
selecting browser-owned canvases. This isolates native Canvas drawing and
small bridge commands from large pixel transfer. Combine it with `--queued`
to compare the completed batch cost, not just enqueue speed.

`--binary` sends uncompressed RGBA bytes in the command response instead of a
JSON array of numbers. The response contains a little-endian four-byte JSON
header length, ordered command metadata, and raw pixel byte ranges. The browser
fetches an ArrayBuffer and passes a Uint8ClampedArray view to standard ImageData
and putImageData. It is not an encoded image and does not require compression,
base64, or a second resource request. Other commands keep their JSON transport.

## Local Debug Measurements

Visible Chrome, short local runs, standard JSON transport:

| Mode | Frames | Submitted Frames Per Second | Main Measured Cost |
|---|---:|---:|---|
| Per-frame upload, synchronous | 16 | 6.19 frames/s | 140.6 ms upload; 8.7 ms two copies |
| Browser-cached, synchronous | 120 | 558.76 frames/s | 1.78 ms two copies |
| Browser-cached, queued + barrier | 120 | 890.91 frames/s | 1.11 ms completion barrier |
| Raw binary upload, synchronous | 120 | 141.63 frames/s | 2.62 ms upload; 4.42 ms two copies |
| Raw binary upload, queued + barrier | 120 | 268.52 frames/s | 3.70 ms completion barrier |
| JSON baseline rerun | 16 | 6.54 frames/s | 135.11 ms upload; 7.13 ms two copies |

The rates count frame submissions, each containing two canvas draws; they do
not count individual draw calls. These runs are not a claim of displayed FPS.
The final pixel readback occurs after timing, so GPU completion/display refresh
is not measured per frame. Preparation and JSON upload dominate the upload
case even without a WASM VM. Different frame counts and warm/cold browser
conditions mean these rows diagnose costs, not a controlled speedup ratio.

Safari and Firefox comparison runs have not produced timing results: both
loaded the HTML shell but failed to connect to the script-driven bridge before
the timeout. Firefox's resource timeline contained no requests for either
bridge script. This is a startup/compatibility gate, not a rendering benchmark
result; do not count blank tabs as zero FPS.
