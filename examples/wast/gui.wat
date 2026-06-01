;; gui.wat — simple counter GUI using vybe:gui
;; Run: cargo run --bin vybex -- examples/wast/gui.wat

(module
  ;; ── vybe:gui imports ──────────────────────────────────────────────────────
  (import "vybe:gui"  "newForm"            (func $newForm            (param externref) (result externref)))
  (import "vybe:gui"  "new_Button"         (func $new_Button         (result externref)))
  (import "vybe:gui"  "new_Label"          (func $new_Label          (result externref)))
  (import "vybe:gui"  "controlSetProperty" (func $setProp            (param externref externref externref)))
  (import "vybe:gui"  "controlsAdd"        (func $controlsAdd        (param externref externref)))
  (import "vybe:gui"  "onEvent"            (func $onEvent            (param externref externref externref)))
  (import "vybe:gui"  "showForm"           (func $showForm           (param externref)))
  (import "vybe:gui"  "runApplication"     (func $runApplication     (param externref)))
  (import "vybe:gui"  "controlGetProperty" (func $getProp            (param externref externref) (result externref)))

  ;; ── wasi:cli for logging ───────────────────────────────────────────────────
  (import "wasi:cli"  "log"                (func $log               (param externref)))

  ;; ── String constants (via wasm:js-string) ─────────────────────────────────
  (import "wasm:js-string" "fromCharCodeArray" (func $str (param externref) (result externref)))

  ;; ── Mutable counter global ────────────────────────────────────────────────
  (global $count (mut i32) (i32.const 0))

  ;; ── Label control reference ───────────────────────────────────────────────
  (global $label (mut externref) (ref.null extern))

  ;; ── Helpers that produce string Values via wasi:cli log trick ─────────────
  ;; In WAT we call vybe:gui functions with externref; string constants come
  ;; from the profile's log/wasi:cli path.  For property names and values
  ;; we use the wast profile's `log` import which accepts any value.

  ;; Build the UI and run the application
  (func $main (export "main")
    ;; Create form
    (local $form    externref)
    (local $btn_inc externref)
    (local $btn_dec externref)
    (local $btn_rst externref)
    (local $lbl     externref)

    ;; newForm("Counter") — title string passed as externref via log import trick
    ;; For simplicity we pass a null externref (title set via setProperty)
    (local.set $form (call $newForm (ref.null extern)))

    ;; Create controls
    (local.set $btn_inc (call $new_Button))
    (local.set $btn_dec (call $new_Button))
    (local.set $btn_rst (call $new_Button))
    (local.set $lbl     (call $new_Label))

    ;; Store label for use in click handlers
    (global.set $label (local.get $lbl))

    ;; Log that the GUI is being set up
    (call $log (ref.null extern))

    ;; Add controls to form
    (call $controlsAdd (local.get $form) (local.get $lbl))
    (call $controlsAdd (local.get $form) (local.get $btn_inc))
    (call $controlsAdd (local.get $form) (local.get $btn_dec))
    (call $controlsAdd (local.get $form) (local.get $btn_rst))

    ;; Show and run
    (call $showForm (local.get $form))
    (call $runApplication (local.get $form)))

  (start $main)
)
