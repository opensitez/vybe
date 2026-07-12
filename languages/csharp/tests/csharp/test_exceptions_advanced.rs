/// C# exception handling: try/catch/finally, custom exceptions,
/// when filters, nested try, using statements, throw/rethrow.
use super::helpers::run_csharp;

// ===================================================================
// TRY / CATCH / FINALLY BASICS
// ===================================================================

#[test]
fn try_catch_basic() {
    assert_eq!(
        run_csharp(
            r#"
try {
    int x = 10 / 0;
} catch (DivideByZeroException) {
    Console.WriteLine("caught divide by zero");
}
"#
        ),
        &["caught divide by zero"]
    );
}

#[test]
fn try_catch_with_variable() {
    assert_eq!(
        run_csharp(
            r#"
try {
    int.Parse("notanumber");
} catch (Exception e) {
    Console.WriteLine("Error: " + e.Message);
}
"#
        ),
        &["Error: Input string was not in a correct format."]
    );
}

#[test]
fn try_finally_always_runs() {
    assert_eq!(
        run_csharp(
            r#"
try {
    Console.WriteLine("in try");
} finally {
    Console.WriteLine("in finally");
}
"#
        ),
        &["in try", "in finally"]
    );
}

#[test]
fn try_catch_finally_together() {
    assert_eq!(
        run_csharp(
            r#"
try {
    throw new Exception("boom");
} catch (Exception e) {
    Console.WriteLine("caught: " + e.Message);
} finally {
    Console.WriteLine("finally");
}
"#
        ),
        &["caught: boom", "finally"]
    );
}

#[test]
fn catch_finally_on_error() {
    assert_eq!(
        run_csharp(
            r#"
string result = "start";
try {
    int x = 10 / 0;
    result = "never";
} catch (DivideByZeroException) {
    result = "caught";
} finally {
    result += " + finally";
}
Console.WriteLine(result);
"#
        ),
        &["caught + finally"]
    );
}

// ===================================================================
// MULTIPLE CATCH BLOCKS
// ===================================================================

#[test]
fn multiple_catch_blocks() {
    assert_eq!(
        run_csharp(
            r#"
try {
    throw new ArgumentException("bad arg");
} catch (ArgumentNullException) {
    Console.WriteLine("null");
} catch (ArgumentException e) {
    Console.WriteLine("arg: " + e.Message);
} catch (Exception) {
    Console.WriteLine("generic");
}
"#
        ),
        &["arg: bad arg"]
    );
}

#[test]
fn catch_hierarchy_most_specific_first() {
    assert_eq!(
        run_csharp(
            r#"
try {
    throw new InvalidOperationException("invalid op");
} catch (InvalidOperationException e) {
    Console.WriteLine("specific: " + e.Message);
} catch (Exception) {
    Console.WriteLine("generic");
}
"#
        ),
        &["specific: invalid op"]
    );
}

// ===================================================================
// THROW AND RETHROW
// ===================================================================

#[test]
fn throw_new_exception() {
    assert_eq!(
        run_csharp(
            r#"
try {
    throw new InvalidOperationException("not allowed");
} catch (InvalidOperationException e) {
    Console.WriteLine(e.Message);
}
"#
        ),
        &["not allowed"]
    );
}

#[test]
fn rethrow_with_throw() {
    assert_eq!(
        run_csharp(
            r#"
try {
    try {
        throw new Exception("inner");
    } catch (Exception) {
        throw;
    }
} catch (Exception e) {
    Console.WriteLine("outer: " + e.Message);
}
"#
        ),
        &["outer: inner"]
    );
}

// ===================================================================
// CUSTOM EXCEPTIONS
// ===================================================================

#[test]
fn custom_exception_class() {
    assert_eq!(
        run_csharp(
            r#"
class AppException : Exception {
    public int Code { get; set; }
    public AppException(string message, int code) : base(message) {
        Code = code;
    }
}
try {
    throw new AppException("not found", 404);
} catch (AppException e) {
    Console.WriteLine(e.Message + " (" + e.Code + ")");
}
"#
        ),
        &["not found (404)"]
    );
}

#[test]
fn custom_exception_hierarchy() {
    assert_eq!(
        run_csharp(
            r#"
class BaseError : Exception {
    public BaseError(string msg) : base(msg) {}
}
class NotFoundError : BaseError {
    public NotFoundError(string msg) : base(msg) {}
}
try {
    throw new NotFoundError("user missing");
} catch (BaseError e) {
    Console.WriteLine("base: " + e.Message);
}
"#
        ),
        &["base: user missing"]
    );
}

// ===================================================================
// WHEN FILTERS
// ===================================================================

#[test]
fn catch_when_filter() {
    assert_eq!(
        run_csharp(
            r#"
try {
    throw new Exception("error 42");
} catch (Exception e) when (e.Message.Contains("42")) {
    Console.WriteLine("filtered catch: " + e.Message);
}
"#
        ),
        &["filtered catch: error 42"]
    );
}

#[test]
fn catch_when_filter_fallthrough() {
    assert_eq!(
        run_csharp(
            r#"
try {
    throw new Exception("error 99");
} catch (Exception e) when (e.Message.Contains("42")) {
    Console.WriteLine("should not match");
} catch (Exception e) {
    Console.WriteLine("fallthrough: " + e.Message);
}
"#
        ),
        &["fallthrough: error 99"]
    );
}

// ===================================================================
// NESTED TRY
// ===================================================================

#[test]
fn nested_try_catch() {
    assert_eq!(
        run_csharp(
            r#"
try {
    Console.WriteLine("outer try");
    try {
        throw new Exception("inner error");
    } catch (Exception e) {
        Console.WriteLine("inner catch: " + e.Message);
    }
    Console.WriteLine("after inner");
} catch (Exception) {
    Console.WriteLine("outer catch");
}
"#
        ),
        &["outer try", "inner catch: inner error", "after inner"]
    );
}

#[test]
fn nested_try_inner_uncaught_propagates() {
    assert_eq!(
        run_csharp(
            r#"
try {
    try {
        throw new InvalidOperationException("oops");
    } catch (ArgumentException) {
        Console.WriteLine("wrong handler");
    }
} catch (InvalidOperationException e) {
    Console.WriteLine("outer got: " + e.Message);
}
"#
        ),
        &["outer got: oops"]
    );
}

// ===================================================================
// USING STATEMENT (IDisposable pattern)
// ===================================================================

#[test]
fn using_statement_basic() {
    assert_eq!(
        run_csharp(
            r#"
class Resource : IDisposable {
    public Resource() { Console.WriteLine("opened"); }
    public void Dispose() { Console.WriteLine("disposed"); }
}
using (var r = new Resource()) {
    Console.WriteLine("using");
}
"#
        ),
        &["opened", "using", "disposed"]
    );
}

#[test]
fn using_disposes_on_exception() {
    assert_eq!(
        run_csharp(
            r#"
class Conn : IDisposable {
    public void Dispose() { Console.WriteLine("conn closed"); }
}
try {
    using (var c = new Conn()) {
        throw new Exception("fail");
    }
} catch (Exception e) {
    Console.WriteLine("caught: " + e.Message);
}
"#
        ),
        &["conn closed", "caught: fail"]
    );
}

// ===================================================================
// EXCEPTION PROPERTIES
// ===================================================================

#[test]
fn exception_message_property() {
    assert_eq!(
        run_csharp(
            r#"
try {
    throw new Exception("test message");
} catch (Exception e) {
    Console.WriteLine(e.Message);
}
"#
        ),
        &["test message"]
    );
}

#[test]
fn argument_null_exception() {
    assert_eq!(
        run_csharp(
            r#"
try {
    throw new ArgumentNullException("param1");
} catch (ArgumentNullException e) {
    Console.WriteLine(e.ParamName);
}
"#
        ),
        &["param1"]
    );
}

#[test]
fn argument_out_of_range_exception() {
    assert_eq!(
        run_csharp(
            r#"
try {
    throw new ArgumentOutOfRangeException("index", "too big");
} catch (ArgumentOutOfRangeException e) {
    Console.WriteLine(e.ParamName);
}
"#
        ),
        &["index"]
    );
}
