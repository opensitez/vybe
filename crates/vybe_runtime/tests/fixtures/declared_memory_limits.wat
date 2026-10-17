(module
  (memory 1 1)
  (func (export "_start")
    (if (i32.ne (memory.size) (i32.const 1)) (then unreachable))
    (if (i32.ne (memory.grow (i32.const 0)) (i32.const 1)) (then unreachable))
    (if (i32.ne (memory.grow (i32.const 1)) (i32.const -1)) (then unreachable))
    (if (i32.ne (memory.size) (i32.const 1)) (then unreachable))
    (if (i32.load8_u (i32.const 65535)) (then unreachable))))
