// vybe-test: kotlin/kotlin_path_and_files/test_files_size_and_exists_checks
// origin: languages/kotlin/tests/kotlin/test_kotlin_path_and_files.rs

import java.nio.file.Files

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val path = Files.createTempFile("vybe-size", ".txt")
            Files.writeString(path, "kotlin")
            __check((Files.exists(path)).toString(), "true")
            __check((Files.isRegularFile(path)).toString(), "true")
            __check((Files.size(path)).toString(), "6")
            Files.delete(path)
            __check((Files.exists(path)).toString(), "false")
        }
