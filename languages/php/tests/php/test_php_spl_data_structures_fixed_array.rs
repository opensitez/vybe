use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: SPL Data Structures — SplFixedArray, SplStack, SplQueue, SplPriorityQueue, SplObjectStorage, ArrayAccess, Countable
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_spl_fixed_array_bounds_and_size() {
    let out = run_prints(
        r#"<?php
$arr = new SplFixedArray(3);
$arr[0] = 10;
$arr[1] = 20;
$arr[2] = 30;
echo count($arr) . " | " . $arr[1];
"#,
    );
    assert_eq!(out, vec!["3 | 20"]);
}

#[test]
fn test_php_spl_stack_push_pop_lifo() {
    let out = run_prints(
        r#"<?php
$stack = new SplStack();
$stack->push("first");
$stack->push("second");
echo $stack->pop() . " -> " . $stack->pop();
"#,
    );
    assert_eq!(out, vec!["second -> first"]);
}

#[test]
fn test_php_spl_queue_enqueue_dequeue_fifo() {
    let out = run_prints(
        r#"<?php
$queue = new SplQueue();
$queue->enqueue("A");
$queue->enqueue("B");
echo $queue->dequeue() . " -> " . $queue->dequeue();
"#,
    );
    assert_eq!(out, vec!["A -> B"]);
}

#[test]
fn test_php_spl_object_storage_attach_detach_contains() {
    let out = run_prints(
        r#"<?php
$storage = new SplObjectStorage();
$o1 = new stdClass();
$o2 = new stdClass();

$storage->attach($o1, "metadata_o1");
echo $storage->contains($o1) ? "YES" : "NO";
echo " ";
echo $storage->contains($o2) ? "YES" : "NO";
"#,
    );
    assert_eq!(out, vec!["YES NO"]);
}

#[test]
fn test_php_array_access_and_countable_interface() {
    let out = run_prints(
        r#"<?php
class ConfigBag implements ArrayAccess, Countable {
    private array $data = [];
    public function offsetExists(mixed $offset): bool { return isset($this->data[$offset]); }
    public function offsetGet(mixed $offset): mixed { return $this->data[$offset] ?? null; }
    public function offsetSet(mixed $offset, mixed $value): void { $this->data[$offset] = $value; }
    public function offsetUnset(mixed $offset): void { unset($this->data[$offset]); }
    public function count(): int { return count($this->data); }
}

$bag = new ConfigBag();
$bag["theme"] = "dark";
$bag["lang"] = "en";
echo count($bag) . " | " . $bag["theme"];
"#,
    );
    assert_eq!(out, vec!["2 | dark"]);
}

#[test]
fn test_php_spl_priority_queue_ordering() {
    compile_ok(
        r#"<?php
$pq = new SplPriorityQueue();
$pq->insert("low priority task", 1);
$pq->insert("high priority task", 100);
$pq->insert("medium priority task", 50);

while ($pq->valid()) {
    echo $pq->current() . "\n";
    $pq->next();
}
"#,
    );
}

#[test]
fn test_php_spl_doubly_linked_list_traversal() {
    compile_ok(
        r#"<?php
$dll = new SplDoublyLinkedList();
$dll->push(1);
$dll->push(2);
$dll->unshift(0);

$dll->setIteratorMode(SplDoublyLinkedList::IT_MODE_FIFO);
foreach ($dll as $val) {
    echo $val . "\n";
}
"#,
    );
}

#[test]
fn test_php_spl_fixed_array_from_array_conversion() {
    compile_ok(
        r#"<?php
$native = [100, 200, 300];
$fixed = SplFixedArray::fromArray($native);
echo $fixed->getSize();
print_r($fixed->toArray());
"#,
    );
}

#[test]
fn test_php_spl_object_storage_associated_data() {
    compile_ok(
        r#"<?php
$storage = new SplObjectStorage();
$user = new stdClass();
$storage[$user] = ["role" => "admin"];

echo $storage[$user]["role"];
"#,
    );
}

#[test]
fn test_php_spl_fixed_array_resize() {
    compile_ok(
        r#"<?php
$fa = new SplFixedArray(2);
$fa[0] = "a";
$fa->setSize(4);
$fa[3] = "d";
echo $fa->getSize() . " | " . $fa[3];
"#,
    );
}
