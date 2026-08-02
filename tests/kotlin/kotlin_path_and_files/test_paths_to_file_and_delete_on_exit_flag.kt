// vybe-test: kotlin/kotlin_path_and_files/test_paths_to_file_and_delete_on_exit_flag
// origin: languages/kotlin/tests/kotlin/test_kotlin_path_and_files.rs

import java.nio.file.Files

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val path = Files.createTempFile("vybe-exit", ".txt")
            path.toFile().deleteOnExit()
            __check((path.toFile().exists()).toString(), "true")
            __check((path.toFile().canWrite()).toString(), "true")
        }
