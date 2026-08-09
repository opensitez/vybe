use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `SuppressedError` & Explicit Resource Management (`using` / `Symbol.dispose` ES2024)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_suppressed_error_constructor_properties() {
    let src = r#"
const primary = new Error("Primary Failure");
const suppressed = new Error("Cleanup Failure");
const err = new SuppressedError(primary, suppressed, "Resource Failure");

console.log(err.name + "|" + err.message + "|" + (err.error === primary) + "|" + (err.suppressed === suppressed));
"#;
    assert_eq!(
        run_js(src),
        vec!["SuppressedError|Resource Failure|true|true"]
    );
}

#[test]
fn test_js_symbol_dispose_well_known_symbol_exists() {
    let src = r#"
console.log(typeof Symbol.dispose === "symbol");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_symbol_async_dispose_well_known_symbol_exists() {
    let src = r#"
console.log(typeof Symbol.asyncDispose === "symbol");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_suppressed_error_instanceof_error() {
    let src = r#"
const err = new SuppressedError(1, 2);
console.log((err instanceof SuppressedError) + "|" + (err instanceof Error));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_suppressed_error_prototype_name() {
    let src = r#"
console.log(SuppressedError.prototype.name);
"#;
    assert_eq!(run_js(src), vec!["SuppressedError"]);
}

#[test]
fn test_js_suppressed_error_factory_call_without_new() {
    let src = r#"
const err = SuppressedError("err1", "err2", "msg");
console.log(err.message + "|" + (err instanceof SuppressedError));
"#;
    assert_eq!(run_js(src), vec!["msg|true"]);
}

#[test]
fn test_js_suppressed_error_primitive_error_and_suppressed() {
    let src = r#"
const err = new SuppressedError(404, "CleanupFailed");
console.log(err.error + "|" + err.suppressed);
"#;
    assert_eq!(run_js(src), vec!["404|CleanupFailed"]);
}

#[test]
fn test_js_symbol_dispose_custom_resource_cleanup() {
    let src = r#"
let disposed = false;
const resource = {
    [Symbol.dispose]() {
        disposed = true;
    }
};
resource[Symbol.dispose]();
console.log(disposed);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_symbol_async_dispose_custom_resource_cleanup() {
    let src = r#"
let asyncDisposed = false;
const resource = {
    async [Symbol.asyncDispose]() {
        await Promise.resolve();
        asyncDisposed = true;
    }
};
(async () => {
    await resource[Symbol.asyncDispose]();
    console.log(asyncDisposed);
})();
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_suppressed_error_cause_support() {
    let src = r#"
const cause = new Error("RootCause");
const err = new SuppressedError(1, 2, "Msg", { cause });
console.log(err.cause.message);
"#;
    assert_eq!(run_js(src), vec!["RootCause"]);
}

#[test]
fn test_js_suppressed_error_tostring_formatting() {
    let src = r#"
const err = new SuppressedError("e1", "e2", "SuppressedDetails");
console.log(err.toString());
"#;
    assert_eq!(run_js(src), vec!["SuppressedError: SuppressedDetails"]);
}

#[test]
fn test_js_suppressed_error_empty_message_defaults_to_empty_string() {
    let src = r#"
const err = new SuppressedError(1, 2);
console.log(err.message === "");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_suppressed_error_nesting() {
    let src = r#"
const e1 = new Error("E1");
const e2 = new Error("E2");
const e3 = new Error("E3");
const inner = new SuppressedError(e1, e2);
const outer = new SuppressedError(inner, e3);

console.log(outer.error.error.message + " -> " + outer.error.suppressed.message);
"#;
    assert_eq!(run_js(src), vec!["E1 -> E2"]);
}

#[test]
fn test_js_disposable_stack_custom_implementation_simulation() {
    let src = r#"
class DisposableStack {
    #resources = [];
    use(res) {
        this.#resources.push(res);
        return res;
    }
    dispose() {
        while (this.#resources.length > 0) {
            const res = this.#resources.pop();
            if (res && typeof res[Symbol.dispose] === "function") {
                res[Symbol.dispose]();
            }
        }
    }
}
const log = [];
const stack = new DisposableStack();
stack.use({ [Symbol.dispose]() { log.push("R1"); } });
stack.use({ [Symbol.dispose]() { log.push("R2"); } });
stack.dispose();
console.log(log.join(",")); // Disposed in LIFO order!
"#;
    assert_eq!(run_js(src), vec!["R2,R1"]);
}

#[test]
fn test_js_async_disposable_stack_custom_implementation_simulation() {
    let src = r#"
class AsyncDisposableStack {
    #resources = [];
    use(res) { this.#resources.push(res); return res; }
    async disposeAsync() {
        while (this.#resources.length > 0) {
            const res = this.#resources.pop();
            if (res && typeof res[Symbol.asyncDispose] === "function") {
                await res[Symbol.asyncDispose]();
            }
        }
    }
}
const log = [];
const stack = new AsyncDisposableStack();
stack.use({ async [Symbol.asyncDispose]() { log.push("AR1"); } });
(async () => {
    await stack.disposeAsync();
    console.log(log.join(","));
})();
"#;
    assert_eq!(run_js(src), vec!["AR1"]);
}

#[test]
fn test_js_suppressed_error_stack_trace() {
    let src = r#"
const err = new SuppressedError("e1", "e2", "StackTest");
console.log(err.stack.includes("SuppressedError: StackTest"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_symbol_dispose_method_property_descriptor() {
    let src = r#"
const desc = Object.getOwnPropertyDescriptor(Symbol, "dispose");
console.log(desc.writable + "|" + desc.enumerable + "|" + desc.configurable);
"#;
    assert_eq!(run_js(src), vec!["false|false|false"]);
}

#[test]
fn test_js_symbol_async_dispose_method_property_descriptor() {
    let src = r#"
const desc = Object.getOwnPropertyDescriptor(Symbol, "asyncDispose");
console.log(desc.writable + "|" + desc.enumerable + "|" + desc.configurable);
"#;
    assert_eq!(run_js(src), vec!["false|false|false"]);
}

#[test]
fn test_js_suppressed_error_null_and_undefined_arguments() {
    let src = r#"
const err = new SuppressedError(null, undefined);
console.log((err.error === null) + "|" + (err.suppressed === undefined));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_suppressed_error_class_inheritance() {
    let src = r#"
class ResourceError extends SuppressedError {}
const err = new ResourceError("a", "b", "CustomResourceError");
console.log(err.name + "|" + err.message + "|isSuppressed=" + (err instanceof SuppressedError));
"#;
    assert_eq!(
        run_js(src),
        vec!["ResourceError|CustomResourceError|isSuppressed=true"]
    );
}

#[test]
fn test_js_suppressed_error_prototype_parent_is_error_prototype() {
    let src = r#"
console.log(Object.getPrototypeOf(SuppressedError.prototype) === Error.prototype);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}
