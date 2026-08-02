<?php
// vybe-test: php/patterns/two_step_view_gather_then_render
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

class ViewModel {
    public array $data = [];
    public function set(string $k, $v): void { $this->data[$k] = $v; }
}
function gatherData(): ViewModel {
    $vm = new ViewModel();
    $vm->set('title', 'My Page');
    $vm->set('items', ['a', 'b', 'c']);
    return $vm;
}
function renderView(ViewModel $vm): string {
    $html = '<h1>' . $vm->data['title'] . '</h1>';
    $html .= '<ul>';
    foreach ($vm->data['items'] as $item) {
        $html .= '<li>' . $item . '</li>';
    }
    $html .= '</ul>';
    return $html;
}
$vm = gatherData();
echo $vm->data['title'];
echo count($vm->data['items']);
echo renderView($vm);

__vybe_check(ob_get_clean(), "My Page3<h1>My Page</h1><ul><li>a</li><li>b</li><li>c</li></ul>");
