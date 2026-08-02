<?php
// vybe-test: php/patterns/facade_simple_interface
// origin: languages/php/tests/php/test_patterns.rs

function __vybe_check($got, $want) {
    // Match the Rust harness's normalisation: strip \r, then drop trailing
    // newlines (it split on "\n" and popped empty trailing elements).
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    // Replay the program's own output so running the file by hand still
    // behaves like the program it was extracted from.
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

class VideoDecoder {
    public function decode(string $f): string { return 'decoded:' . $f; }
}
class AudioMixer {
    public function mix(string $a): string { return 'mixed:' . $a; }
}
class VideoFacade {
    private $decoder;
    private $mixer;
    public function __construct() {
        $this->decoder = new VideoDecoder();
        $this->mixer = new AudioMixer();
    }
    public function process(string $file): string {
        $v = $this->decoder->decode($file);
        $a = $this->mixer->mix($file);
        return $v . '|' . $a;
    }
}
echo (new VideoFacade())->process('movie.mp4');

__vybe_check(ob_get_clean(), "decoded:movie.mp4|mixed:movie.mp4");
