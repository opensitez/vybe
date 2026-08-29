;; vybe-test: wast/wat_js_prototypes/configure_all_accepts_an_empty_configuration
;; origin: proposals/custom-descriptors/.../Overview.md §"Configuration API"

;; A stream of a single zero byte is `vec(protoconfig)` with no elements — a
;; well-formed configuration that installs nothing. Null `prototypes` and
;; `functions` are legal here: the params are `(ref null …)` and nothing is
;; consumed.
(module
  (type $prototypes (array (mut externref)))
  (type $functions (array (mut funcref)))
  (type $data (array (mut i8)))
  (type $configureAll (func (param (ref null $prototypes))
                            (param (ref null $functions))
                            (param (ref null $data))
                            (param externref)))
  (import "wasm:js-prototypes" "configureAll" (func $configureAll (type $configureAll)))
  (func (export "_start")
    (call $configureAll
      (ref.null $prototypes)
      (ref.null $functions)
      (array.new_fixed $data 1 (i32.const 0))
      (ref.null extern))))
