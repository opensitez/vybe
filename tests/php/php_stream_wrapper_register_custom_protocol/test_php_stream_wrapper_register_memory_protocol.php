<?php
// vybe-test: php/php_stream_wrapper_register_custom_protocol/test_php_stream_wrapper_register_memory_protocol
// origin: languages/php/tests/php/test_php_stream_wrapper_register_custom_protocol.rs

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

class MemoryStreamWrapper {
    public $context;
    private string $data = "";
    private int $position = 0;

    public function stream_open(string $path, string $mode, int $options, ?string &$opened_path): bool {
        $this->position = 0;
        return true;
    }

    public function stream_write(string $data): int {
        $this->data .= $data;
        $this->position += strlen($data);
        return strlen($data);
    }

    public function stream_read(int $count): string {
        $ret = substr($this->data, $this->position, $count);
        $this->position += strlen($ret);
        return $ret;
    }

    public function stream_tell(): int { return $this->position; }
    public function stream_eof(): bool { return $this->position >= strlen($this->data); }
}

stream_wrapper_register("memoryvar", MemoryStreamWrapper::class);

$fp = fopen("memoryvar://test", "r+");
fwrite($fp, "Hello Custom Stream");
rewind($fp);

$read = stream_get_contents($fp);
fclose($fp);
stream_wrapper_unregister("memoryvar");

echo $read;

__vybe_check(ob_get_clean(), "Hello Custom Stream");
