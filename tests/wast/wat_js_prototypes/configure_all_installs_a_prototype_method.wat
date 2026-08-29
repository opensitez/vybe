;; vybe-test: wast/wat_js_prototypes/configure_all_installs_a_prototype_method
;; origin: proposals/custom-descriptors/.../Overview.md §"Configuration API"

;; `wasm:js-prototypes.configureAll` parses the configuration byte stream and
;; consumes the prototypes and functions arrays IN ORDER: one entry of
;; `prototypes` per `protoconfig`, one entry of `functions` per
;; `constructorconfig` or `methodconfig`.
;;
;; Stream below: 1 protoconfig / 0 constructorconfigs / 1 methodconfig
;; (kind 0x00 = method, name "get") / parentidx -1 (0x7F as signed LEB).
(module
  (type $prototypes (array (mut externref)))
  (type $functions (array (mut funcref)))
  (type $data (array (mut i8)))
  (type $s (struct (field i32)))
  (type $ft (func))
  (type $configureAll (func (param (ref null $prototypes))
                            (param (ref null $functions))
                            (param (ref null $data))
                            (param externref)))
  (import "wasm:js-prototypes" "configureAll" (func $configureAll (type $configureAll)))
  (elem declare func $m)
  (func $m (type $ft))
  (func (export "_start")
    (call $configureAll
      (array.new_fixed $prototypes 1 (extern.convert_any (struct.new $s (i32.const 7))))
      (array.new_fixed $functions 1 (ref.func $m))
      (array.new_fixed $data 9
        (i32.const 1) (i32.const 0) (i32.const 1) (i32.const 0)
        (i32.const 3) (i32.const 103) (i32.const 101) (i32.const 116)
        (i32.const 127))
      (ref.null extern))))
