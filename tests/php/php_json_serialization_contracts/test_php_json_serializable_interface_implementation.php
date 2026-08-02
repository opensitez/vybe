<?php
// vybe-test: php/php_json_serialization_contracts/test_php_json_serializable_interface_implementation
// origin: languages/php/tests/php/test_php_json_serialization_contracts.rs

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

class ApiResponse implements JsonSerializable {
    public function __construct(
        private string $status,
        private array $data
    ) {}

    public function jsonSerialize(): mixed {
        return [
            "success" => $this->status === "ok",
            "payload" => $this->data,
        ];
    }
}

$res = new ApiResponse("ok", ["user_id" => 42]);
echo json_encode($res);

__vybe_check(ob_get_clean(), "{\"success\":true,\"payload\":{\"user_id\":42}}");
