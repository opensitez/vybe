<?php
// vybe-test: interop/destructor_slot/test_php_destroys_pascal_object
// vybe-test-units: lib_pascal.pas
//
// THE test the protocol-slot architecture exists for, and the one the corpus
// never had. `documentation/flexclassplan.md` §0.2 — "identity is structural,
// never lexical" — is a claim about exactly this: PHP must destroy a Pascal
// object without knowing that Pascal spells its destructor `Destroy`.
//
// `classes.rs` binds every role method under both its source name AND
// `__vybe_slot_<id>`, so `unset($r)` has something to reach for. Publication
// has been in place for months and nothing consumed it — and nothing noticed,
// because no test in the corpus crosses a language boundary. That is also how
// `project_java_tostring_slot_unread` happened.
//
// Units link in order with the ENTRY LAST:
//     vybex lib_pascal.pas test_php_destroys_pascal_object.php
// `cli.rs` — "Link the other languages in first ... puts its functions and
// classes in the shared global table and runs its top-level code, so the entry
// unit starts with everything already defined."
//
// EXPECTED TO FAIL when written. PHP's `unset` rewrite asks
// `php_object_class_from_expr` for the class and then `class_has_destructor`,
// both of which consult PHP's OWN class registry — a Pascal class is not in
// it, so the destructor branch is dead and nothing runs. The fix is for that
// path to emit the Destructor slot instead of the name `__destruct`.

function __vybe_check($got, $want) {
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

$r = MakeResource("db-handle");

// Reading a field across the boundary is the easy half.
echo $r->Name, "\n";

// The hard half: PHP's own destruction verb has to run PASCAL's destructor.
// `unset` must lower to the Destructor SLOT, never to the string
// "__destruct" — which a Pascal class does not have and never will.
unset($r);

echo GetDestroyLog(), "\n";

__vybe_check(ob_get_clean(), "db-handle\ndestroyed:db-handle;");
