<?php
// vybe-test: php/patterns/unit_of_work_track_commit
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

class UnitOfWork {
    private $new = [];
    private $dirty = [];
    private $deleted = [];
    public function registerNew(object $entity): void { $this->new[] = $entity; }
    public function registerDirty(object $entity): void { $this->dirty[] = $entity; }
    public function registerDeleted(object $entity): void { $this->deleted[] = $entity; }
    public function commit(): array {
        return [
            'inserted' => count($this->new),
            'updated' => count($this->dirty),
            'deleted' => count($this->deleted),
        ];
    }
}
$uow = new UnitOfWork();
$uow->registerNew((object)['id' => 1]);
$uow->registerNew((object)['id' => 2]);
$uow->registerDirty((object)['id' => 3]);
$uow->registerDeleted((object)['id' => 4]);
$result = $uow->commit();
echo $result['inserted'];
echo $result['updated'];
echo $result['deleted'];

__vybe_check(ob_get_clean(), "211");
