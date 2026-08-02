// vybe-test: kotlin/kotlin_path_and_files/test_path_file_name_and_parent
// origin: languages/kotlin/tests/kotlin/test_kotlin_path_and_files.rs

import java.nio.file.Paths

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val path = Paths.get("/tmp", "alpha", "beta", "data.txt")
            __check((path.fileName.toString()).toString(), "data.txt")
            __check((path.fileName.toString().length).toString(), "8")
            __check((path.parent?.fileName?.toString()).toString(), "beta")
            __check((path.root?.toString()).toString(), "/")
        }
