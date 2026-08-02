<?php
// vybe-test: php/php_attributes_inheritance_traits/parent_class_walk_collects_mapped_superclass_columns
// origin: languages/php/tests/php/test_php_attributes_inheritance_traits.rs

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

#[Attribute]
class Column {
    public function __construct(public string $type) {}
}
class Timestamps {
    #[Column('datetime')]
    public $createdAt;
}
class Post extends Timestamps {
    #[Column('string')]
    public $title;
}
$names = [];
for ($rc = new ReflectionClass(Post::class); $rc; $rc = $rc->getParentClass()) {
    foreach ($rc->getProperties() as $p) {
        foreach ($p->getAttributes(Column::class) as $a) {
            $names[] = $rc->getName() . '.' . $p->getName() . ':' . $a->newInstance()->type;
        }
    }
}
echo implode('|', $names);

__vybe_check(ob_get_clean(), "Post.title:string|Post.createdAt:datetime|Timestamps.createdAt:datetime");
