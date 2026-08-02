// vybe-test: kotlin/kotlin_path_and_files/test_path_resolve_and_normalize
// origin: languages/kotlin/tests/kotlin/test_kotlin_path_and_files.rs

import java.nio.file.Paths

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val root = Paths.get("/tmp", "base")
            val child = root.resolve("a").resolve("../b.txt").normalize()
            __check((child.toString()).toString(), "/tmp/base/b.txt")
            __check((child.endsWith("b.txt")).toString(), "true")
        }
