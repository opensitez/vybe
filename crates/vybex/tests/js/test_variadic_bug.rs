//! Regression tests for variadic function parameter slot allocation.
//!
//! Bug: `function f(fmt, ...args) { ... fmt.charAt(i) ... }` corrupts
//! `args` after the first method call inside the body. Root cause was
//! `emitter::invoke::emit_invoke_method` using `chunk.local_count` as
//! the base for temp slots — but the compiler doesn't sync
//! `chunk.local_count` during compilation, so it stayed at 0 and the
//! temps landed on top of slot 0 (fmt) and slot 1 (args), wiping the
//! parameters mid-function.
//!
//! The fix: every function-chunk producer (compile_lambda /
//! compile_function_decl) must seed `chunk.local_count` to the
//! high-water mark of slots its prologue uses (params + rest-source
//! reserve + __rest_arr) so `emit_invoke_method`'s temp_base lands
//! safely above all named slots.

use crate::helpers::run_js;

fn run_one(src: &str) -> String {
    run_js(src).join(" ")
}

#[test]
fn variadic_args_survive_method_call_in_loop() {
    // Reproduces the original sprintf polyfill bug. Without the fix,
    // `args` becomes empty after the first `fmt.charAt(i)` call.
    let out = run_one(r#"
        function f(fmt, ...args) {
            let out = "";
            let i = 0;
            const len = fmt.length;
            while (i < len) {
                const c = fmt.charAt(i);
                out += c;
                i++;
            }
            return [args.length, args[0]];
        }
        const r = f("=%s=", "X");
        console.log(r[0], r[1]);
    "#);
    // Expected: rest collected `["X"]` survives the loop.
    assert_eq!(out, "1 X");
}

#[test]
fn variadic_args_survive_through_char_at_only() {
    // Single iteration with a single charAt — minimal repro.
    let out = run_one(r#"
        function f(fmt, ...args) {
            const c = fmt.charAt(0);
            return [args.length, args[0], c];
        }
        const r = f("=", "X");
        console.log(r[0], r[1], r[2]);
    "#);
    assert_eq!(out, "1 X =");
}

#[test]
fn variadic_args_after_typed_method_call() {
    // `String.prototype.toUpperCase` is also a polymorphic method
    // dispatch — same code path. Confirm it doesn't clobber rest either.
    let out = run_one(r#"
        function f(fmt, ...args) {
            const u = fmt.toUpperCase();
            return [args.length, args[0], u];
        }
        const r = f("hi", "X");
        console.log(r[0], r[1], r[2]);
    "#);
    assert_eq!(out, "1 X HI");
}

#[test]
fn variadic_args_after_dict_literal() {
    // `emit_array_pair` (collections.rs) historically had the same
    // chunk.local_count-as-scratch bug as emit_invoke_method. The
    // `define_local` wrapper now keeps chunk.local_count synced, so
    // a dict literal inside a variadic function preserves rest args.
    let out = run_one(r#"
        function f(first, ...rest) {
            const obj = { a: 1, b: 2, c: 3 };
            return [rest.length, rest[0], Object.keys(obj).length];
        }
        const r = f("X", "Y", "Z");
        console.log(r[0], r[1], r[2]);
    "#);
    assert_eq!(out, "2 Y 3");
}

#[test]
fn variadic_args_through_method_chain() {
    // Multiple polymorphic dispatches stacked — exercises the cumulative
    // chunk.local_count + scope.next_slot synchronization.
    let out = run_one(r#"
        function f(s, ...args) {
            const r = s.toUpperCase().toLowerCase().trim();
            return [args.length, args[0], r];
        }
        const r = f("  Hi  ", 1, 2);
        console.log(r[0], r[1], r[2]);
    "#);
    assert_eq!(out, "2 1 hi");
}

#[test]
fn variadic_instance_method_call_packs_rest() {
    let out = run_one(r#"
        class Greeter {
            call(prefix, ...parts) {
                return prefix + ":" + parts.join(",");
            }
        }
        const g = new Greeter();
        console.log(g.call("head", "a", "b", "c"));
    "#);
    assert_eq!(out, "head:a,b,c");
}

#[test]
fn variadic_static_method_call_packs_rest() {
    let out = run_one(r#"
        class Greeter {
            static call(prefix, ...parts) {
                return prefix + ":" + parts.join(",");
            }
        }
        console.log(Greeter.call("head", "a", "b", "c"));
    "#);
    assert_eq!(out, "head:a,b,c");
}

#[test]
fn variadic_instance_method_alias_packs_rest() {
    let out = run_one(r#"
        class Greeter {
            call(prefix, ...parts) {
                return prefix + ":" + parts.join(",");
            }
        }
        const g = new Greeter();
        const alias = g.call;
        console.log(alias("head", "a", "b", "c"));
    "#);
    assert_eq!(out, "head:a,b,c");
}

#[test]
fn variadic_static_method_alias_packs_rest() {
    let out = run_one(r#"
        class Greeter {
            static call(prefix, ...parts) {
                return prefix + ":" + parts.join(",");
            }
        }
        const alias = Greeter.call;
        console.log(alias("head", "a", "b", "c"));
    "#);
    assert_eq!(out, "head:a,b,c");
}

#[test]
fn variadic_function_alias_packs_rest() {
    let out = run_one(r#"
        function joinWith(prefix, ...parts) {
            return prefix + ":" + parts.join(",");
        }
        const alias = joinWith;
        console.log(alias("head", "a", "b", "c"));
    "#);
    assert_eq!(out, "head:a,b,c");
}
