(module
  (memory 1)
  (func $exercise (param $dst i32) (param $src i32) (param $size i32)
    (local $iteration i32) (local $byte i32)
    local.get $src
    i32.const 37
    local.get $size
    memory.fill
    loop $repeat
      local.get $dst
      local.get $src
      local.get $size
      memory.copy
      local.get $dst
      i32.const 37
      local.get $size
      memory.fill
      local.get $iteration
      i32.const 1
      i32.add
      local.tee $iteration
      i32.const 8192
      i32.lt_u
      br_if $repeat
    end
    loop $check
      local.get $dst
      local.get $byte
      i32.add
      i32.load8_u
      i32.const 37
      i32.ne
      if unreachable end
      local.get $byte
      i32.const 1
      i32.add
      local.tee $byte
      local.get $size
      i32.lt_u
      br_if $check
    end)
  (func (export "_start")
    i32.const 8192 i32.const 0 i32.const 8 call $exercise
    i32.const 8192 i32.const 0 i32.const 64 call $exercise
    i32.const 8192 i32.const 0 i32.const 4096 call $exercise
    i32.const 12288 i32.load8_u
    if unreachable end))
