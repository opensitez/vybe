<?php
// vybe-test: php/programs/trie_structure_insert_search
// origin: languages/php/tests/php/test_programs.rs

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

class TrieNode {
    public array $children = [];
    public bool $end = false;
}
class Trie {
    private TrieNode $root;
    public function __construct() { $this->root = new TrieNode(); }
    public function insert(string $word): void {
        $node = $this->root;
        foreach (str_split($word) as $c) {
            if (!isset($node->children[$c])) $node->children[$c] = new TrieNode();
            $node = $node->children[$c];
        }
        $node->end = true;
    }
    public function search(string $word): bool {
        $node = $this->root;
        foreach (str_split($word) as $c) {
            if (!isset($node->children[$c])) return false;
            $node = $node->children[$c];
        }
        return $node->end;
    }
    public function startsWith(string $prefix): bool {
        $node = $this->root;
        foreach (str_split($prefix) as $c) {
            if (!isset($node->children[$c])) return false;
            $node = $node->children[$c];
        }
        return true;
    }
}
$trie = new Trie();
$trie->insert('apple');
$trie->insert('app');
echo $trie->search('app') ? 'true' : 'false';
echo "\n";
echo $trie->search('apple') ? 'true' : 'false';
echo "\n";
echo $trie->search('ap') ? 'true' : 'false';
echo "\n";
echo $trie->startsWith('ap') ? 'true' : 'false';
echo "\n";

__vybe_check(ob_get_clean(), "true\ntrue\nfalse\ntrue");
