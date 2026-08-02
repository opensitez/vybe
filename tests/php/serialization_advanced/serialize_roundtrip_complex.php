<?php
// vybe-test: php/serialization_advanced/serialize_roundtrip_complex
// origin: languages/php/tests/php/test_serialization_advanced.rs
// vybe-test-mode: compile

class Tree {
    public array $children = [];
    public function __construct(public string $label) {}
    public function addChild(Tree $child): void { $this->children[] = $child; }
}
$root = new Tree('root');
$root->addChild(new Tree('child1'));
$root->addChild(new Tree('child2'));
$root->children[0]->addChild(new Tree('grandchild'));
$s = serialize($root);
$r = unserialize($s);
echo $r->label . ':' . count($r->children) . ':' . $r->children[0]->label;
