// vybe-test: java/nio_files_api/files_mismatch_detects_different_content
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
java.nio.file.Path a = java.nio.file.Files.createTempFile("vybe", ".a"); java.nio.file.Path b = java.nio.file.Files.createTempFile("vybe", ".b"); java.nio.file.Files.writeString(a, "aaa"); java.nio.file.Files.writeString(b, "bbb"); long pos = java.nio.file.Files.mismatch(a, b); __p(pos >= 0); java.nio.file.Files.delete(a); java.nio.file.Files.delete(b);
__check("true");
    }
}

