<?php
// vybe-test: interop/tostring_slot/test_php_stringifies_pascal_object
// vybe-test-units: lib_pascal.pas
//
// The CONTROL for the destructor pair next door. `ToString` is a slot the
// shared compiler already consumes — `expressions.rs:7824` emits
// `protocol_slot_key(ProtocolSlot::ToString)` — so if this passes and the
// destructor one fails, the difference is the CALL SITE, not the slot
// machinery. Two tests that fail together mean something more basic (unit
// linking, class visibility across units) is wrong, and that is worth being
// able to tell apart at a glance.
//
// Pascal spells it `ToString`, PHP spells it `__toString`. String
// interpolation is PHP asking for the coercion role; nothing in this file
// names Pascal's spelling.

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

$p = MakePoint(3, 4);

// Field reads across the boundary.
echo $p->X + $p->Y, "\n";

// The role: PHP's string coercion must reach Pascal's `ToString`.
echo "point=" . $p . "\n";

__vybe_check(ob_get_clean(), "7\npoint=(3,4)");
