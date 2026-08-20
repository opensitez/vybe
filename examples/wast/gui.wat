;; gui.wat — a working counter GUI on the WEB platform.
;; Run: cargo run --bin vybex -- examples/wast/gui.wat
;;
;; A label shows the count; the +, -, and Reset buttons update a mutable global
;; and refresh the label. There is no toolkit API here: a
;; control IS a DOM element, so this builds `<button>`/`<div>` nodes with
;; `createElement`, puts them in the document with `appendChild`, and binds
;; clicks with `addEventListener` — the same calls a browser would answer.
;;
;; Nothing tells the page to run. It runs because it HAS a document, which is
;; also what makes the window open (`gui_document::has_content`).

(module
  ;; ── web:dom — Node/Document members ────────────────────────────────────────
  (import "web:dom" "createElement"
    (func $createElement (param externref externref externref) (result externref))) ;; doc,tag,type -> node
  (import "web:dom" "appendChild"
    (func $appendChild (param externref externref externref) (result externref)))   ;; doc,parent,child -> node
  (import "web:dom" "setTextContent"
    (func $setTextContent (param externref externref externref)))                          ;; doc,node,text
  (import "web:dom" "addEventListener"
    (func $addEventListener (param externref externref externref funcref)))         ;; doc,node,type,handler

  ;; ── web:html — HTML element IDL ────────────────────────────────────────────
  (import "web:html" "activeDocument" (func $activeDocument (result externref)))
  (import "web:html" "body"           (func $body (param externref) (result externref)))
  (import "web:html" "setTitle"       (func $setTitle (param externref externref)))

  ;; ── web:cssom — CSSStyleDeclaration ────────────────────────────────────────
  (import "web:cssom" "setStyleProperty"
    (func $setStyleProperty (param externref externref externref externref)))               ;; doc,node,prop,value

  (import "wasi:cli" "log" (func $log (param externref)))

  ;; ── Counter state ──────────────────────────────────────────────────────────
  (global $count (mut i32) (i32.const 0))
  ;; The label element itself, not a name to look one up by: the handle IS the
  ;; identity, so a refresh needs no search.
  (global $label (mut externref) (ref.null extern))

  ;; Repaint the label with the current count.
  (func $refresh
    (call $setTextContent (call $activeDocument) (global.get $label) (global.get $count)))

  ;; ── Button click handlers ──────────────────────────────────────────────────
  (func $inc
    (global.set $count (i32.add (global.get $count) (i32.const 1)))
    (call $refresh))
  (func $dec
    (global.set $count (i32.sub (global.get $count) (i32.const 1)))
    (call $refresh))
  (func $reset
    (global.set $count (i32.const 0))
    (call $refresh))

  ;; Make one `<button>`, caption it, put it in the body, bind its click.
  (func $button (param $caption externref) (param $handler funcref) (result externref)
    (local $node externref)
    (local.set $node
      (call $createElement (call $activeDocument) (string.const "button") (string.const "")))
    (call $setTextContent (call $activeDocument) (local.get $node) (local.get $caption))
    (drop (call $appendChild (call $activeDocument) (call $body (call $activeDocument)) (local.get $node)))
    (call $addEventListener
      (call $activeDocument) (local.get $node) (string.const "click") (local.get $handler))
    (local.get $node))

  ;; ── Build the UI and run ───────────────────────────────────────────────────
  (func $main (export "main")
    (call $setTitle (call $activeDocument) (string.const "Counter"))

    ;; Count label — a block-level element so the buttons sit on the next line.
    (global.set $label
      (call $createElement (call $activeDocument) (string.const "div") (string.const "")))
    (call $setStyleProperty (call $activeDocument) (global.get $label)
      (string.const "font-size") (string.const "24px"))
    (drop (call $appendChild (call $activeDocument) (call $body (call $activeDocument))
      (global.get $label)))
    (call $refresh)

    (drop (call $button (string.const "+") (ref.func $inc)))
    (drop (call $button (string.const "-") (ref.func $dec)))
    (drop (call $button (string.const "Reset") (ref.func $reset)))

    (call $log (string.const "Counter ready - use + / - / Reset")))

  (start $main)
)
