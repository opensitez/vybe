use crate::helpers::run_in_main;

macro_rules! jm {
    ($name:ident, $src:expr, $types:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_in_main($src, $types), vec![$expected]);
        }
    };
}

jm!(
    catch_runtime_exception,
    "try { throw new RuntimeException(\"x\"); } catch (Exception e) { System.out.println(e.getClass().getSimpleName()); }",
    "",
    "RuntimeException"
);
jm!(
    finally_executes_on_success,
    "StringBuilder sb = new StringBuilder(); try { sb.append(\"ok\"); } finally { sb.append(\"+f\"); } System.out.println(sb.toString());",
    "",
    "ok+f"
);
jm!(
    finally_executes_on_exception,
    "StringBuilder sb = new StringBuilder(); try { sb.append(\"a\"); throw new RuntimeException(); } catch (RuntimeException e) { sb.append(\"c\"); } finally { sb.append(\"f\"); } System.out.println(sb.toString());",
    "",
    "acf"
);
jm!(
    nested_try_catch,
    "try { try { int[] values = {1}; System.out.println(values[1]); } catch (IndexOutOfBoundsException e) { System.out.println(\"inner\"); } } catch (Exception e) { System.out.println(\"outer\"); }",
    "",
    "inner"
);
jm!(
    custom_exception_round_trip,
    "class MyErr extends RuntimeException { MyErr(String s) { super(s); } } try { throw new MyErr(\"ok\"); } catch (MyErr e) { System.out.println(e.getMessage()); }",
    "",
    "ok"
);
jm!(
    multi_catch_supported,
    "try { int[] values = {}; System.out.println(values[1]); } catch (NullPointerException | ArrayIndexOutOfBoundsException e) { System.out.println(\"bad\"); }",
    "",
    "bad"
);
jm!(
    rethrow_in_catch,
    "try { try { throw new IllegalStateException(); } catch (IllegalStateException e) { throw e; } } catch (IllegalStateException e) { System.out.println(\"again\"); }",
    "",
    "again"
);
jm!(
    exception_chain_and_order,
    "try { int z = 1 / 0; } catch (ArithmeticException e) { System.out.println(\"arith\"); } catch (RuntimeException e) { System.out.println(\"runtime\"); }",
    "",
    "arith"
);
jm!(
    null_pointer_to_string,
    "try { String s = null; System.out.println(s.length()); } catch (NullPointerException e) { System.out.println(\"np\"); }",
    "",
    "np"
);
jm!(
    finally_always_runs_with_nested,
    "StringBuilder sb = new StringBuilder(); try { try { sb.append(\"a\"); } finally { sb.append(\"b\"); } } finally { sb.append(\"c\"); } System.out.println(sb.toString());",
    "",
    "abc"
);
jm!(
    try_with_checked_exception_method,
    "try { Checked.raise(); } catch (Exception e) { System.out.println(e.getMessage()); }",
    "static class Checked { static void raise() throws Exception { throw new Exception(\"x\"); } }",
    "x"
);
jm!(
    finally_after_return_path,
    "StringBuilder sb = new StringBuilder(); try { try { sb.append(\"A\"); } finally { sb.append(\"B\"); return; } } catch (Exception e) { sb.append(\"C\"); } System.out.println(sb.toString());",
    "",
    "AB"
);
jm!(
    multiple_finally_statements,
    "int value = 0; try { value = 1; } finally { value = 2; } System.out.println(value);",
    "",
    "2"
);
jm!(
    nested_catch_blocks,
    "try { int[] values = {1,2,3}; System.out.println(values[3]); } catch (IndexOutOfBoundsException e) { System.out.println(\"bounds\"); } catch (RuntimeException e) { System.out.println(\"runtime\"); }",
    "",
    "bounds"
);
jm!(
    add_suppressed_exception,
    "try { try { throw new RuntimeException(); } catch (RuntimeException e) { Exception wrapped = new Exception(\"wrapped\"); wrapped.addSuppressed(e); throw wrapped; } } catch (Exception e) { System.out.println(e.getSuppressed().length); }",
    "",
    "1"
);
jm!(
    message_length_on_throw,
    "try { throw new RuntimeException(\"boom\"); } catch (RuntimeException e) { System.out.println(e.getMessage().length()); }",
    "",
    "4"
);
jm!(
    catch_with_renamed_local,
    "int base = 1; try { throw new RuntimeException(); } catch (RuntimeException error) { System.out.println(base + 1); }",
    "",
    "2"
);
jm!(
    rethrow_and_catch_outer,
    "try { try { throw new IllegalArgumentException(); } catch (IllegalArgumentException e) { throw e; } } catch (RuntimeException e) { System.out.println(\"outer\"); }",
    "",
    "outer"
);
jm!(
    finally_after_try_without_exception,
    "StringBuilder sb = new StringBuilder(); try { sb.append(\"ok\"); } finally { sb.append(\"f\"); } System.out.println(sb.toString());",
    "",
    "okf"
);
jm!(
    throw_in_finally_replaces_original,
    "try { try { throw new IllegalArgumentException(); } finally { throw new RuntimeException(); } } catch (RuntimeException e) { System.out.println(\"runtime\"); }",
    "",
    "runtime"
);
jm!(
    nested_finally_and_catch,
    "int x = 1; try { try { x = 2; throw new RuntimeException(); } catch (RuntimeException e) { x = 3; } finally { x = 4; } } catch (RuntimeException e) { System.out.println(\"bad\"); } System.out.println(x);",
    "",
    "4"
);
jm!(
    manual_resource_like_cleanup,
    "java.io.StringReader reader = new java.io.StringReader(\"x\"); try { int c = reader.read(); System.out.println(c); } catch (Exception e) { System.out.println(-1); }",
    "",
    "120"
);
jm!(
    runtime_exception_rethrow_message,
    "try { try { throw new RuntimeException(\"first\"); } catch (RuntimeException e) { throw new RuntimeException(e.getMessage() + \"-x\"); } } catch (RuntimeException e) { System.out.println(e.getMessage()); }",
    "",
    "first-x"
);
