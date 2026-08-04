// vybe-test: java/inheritance_core/super_call_from_overridden_void_method_runs_parent_logic
// origin: languages/java/tests/java/test_inheritance_core.rs

public class Main {
static class Logger { void log(String msg) { System.out.println("parent:" + msg); } }
        static class AuditLogger extends Logger {
            void log(String msg) { super.log(msg); System.out.println("child:" + msg); }
        }
    public static void main(String[] args) {
AuditLogger a = new AuditLogger(); a.log("x");
    }
}

