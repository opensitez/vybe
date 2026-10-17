(module
  (func (export "_start") (local $average v128)
    ;; Keep an unrelated operand below the SIMD operation and its checks.
    i32.const 37
    v128.const i16x8 0 1 2 32767 32768 65534 65535 65535
    v128.const i16x8 65535 0 1 32768 32767 65535 65535 0
    i16x8.avgr_u
    local.tee $average
    v128.const i16x8 32768 1 2 32768 32768 65535 65535 32768
    i16x8.eq
    i16x8.bitmask
    i32.const 255
    i32.ne
    if unreachable end
    i32.const 37
    i32.ne
    if unreachable end

    v128.const i16x8 -32768 32767 -1 0 1 -2 0 -32768
    i16x8.bitmask
    i32.const 165
    i32.ne
    if unreachable end

    local.get $average
    i16x8.bitmask
    i32.const 249
    i32.ne
    if unreachable end))
