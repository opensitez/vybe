// vybe-test: java/nio_files_api/files_read_all_lines_returns_line_list
// origin: languages/java/tests/java/test_nio_files_api.rs

public class Main {

    // A static String, NOT a StringBuilder. Calling a method on a bare static
    // FIELD receiver fails under Vybe with "undefined is not callable"
    // (measured): `SB.append(x)` throws while `StringBuilder l = SB;
    // l.append(x)` works, so the method is resolved from the receiver's
    // declared type at the call site and a static field carries none. String
    // concatenation onto a static field has no such problem.
    static String __buf = "";

    static void __p(Object o) {
        __buf = __buf + String.valueOf(o) + "\n";
    }

    static void __pr(Object o) {
        __buf = __buf + String.valueOf(o);
    }

    static void __check(String want) {
        String got = __buf;
        // The final `println` contributes a trailing newline that the expected
        // line vector never carried, so it is not part of the comparison.
        if (got.endsWith("\n")) {
            got = got.substring(0, got.length() - 1);
        }
        if (!got.equals(want)) {
            System.out.println("FAIL: want [" + want + "] got [" + got + "]");
            throw new RuntimeException("assertion failed");
        }
    }

    public static void main(String[] args) {
java.nio.file.Path p = java.nio.file.Files.createTempFile("vybe", ".txt"); java.nio.file.Files.writeString(p, "line1\nline2"); java.util.List<String> lines = java.nio.file.Files.readAllLines(p); __p(lines.size()); __p(lines.get(1)); java.nio.file.Files.delete(p);
__check("2\nline2");
    }
}

