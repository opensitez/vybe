(module
  (func (export "_start")
    (local $a i32) (local $b i32) (local $wide i64)
    (local $nan f64) (local $reference externref)
    i32.const 17 local.set $a
    i32.const 20 local.set $b

    i32.const 99
    local.get $a local.get $b i32.add
    i32.const 37 i32.ne
    if unreachable end
    i32.const 99 i32.ne
    if unreachable end

    local.get $a local.get $a i32.add
    i32.const 34 i32.ne
    if unreachable end

    ;; Existing get/get/set fusion must retain its priority and ordering.
    local.get $a local.get $b local.set $a
    i32.const 17 i32.ne
    if unreachable end
    local.get $a i32.const 20 i32.ne
    if unreachable end

    i64.const 9223372036854775807 local.set $wide
    local.get $wide local.get $wide i64.xor
    i64.eqz i32.eqz
    if unreachable end

    f64.const nan:0x8000000000042 local.set $nan
    local.get $nan local.get $nan drop
    i64.reinterpret_f64
    i64.const 0x7ff8000000000042 i64.ne
    if unreachable end

    local.get $reference local.get $reference drop
    ref.is_null i32.eqz
    if unreachable end))
