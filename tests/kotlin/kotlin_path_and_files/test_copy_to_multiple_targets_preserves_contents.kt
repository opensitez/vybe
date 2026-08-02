// vybe-test: kotlin/kotlin_path_and_files/test_copy_to_multiple_targets_preserves_contents
// origin: languages/kotlin/tests/kotlin/test_kotlin_path_and_files.rs

import java.nio.file.Files
        import java.nio.file.StandardCopyOption

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val src = Files.createTempFile("vybe-copy-a", ".txt")
            Files.writeString(src, "data")
            val d1 = Files.createTempFile("vybe-copy-b", ".txt")
            val d2 = Files.createTempFile("vybe-copy-c", ".txt")
            Files.copy(src, d1, StandardCopyOption.REPLACE_EXISTING)
            Files.copy(src, d2, StandardCopyOption.REPLACE_EXISTING)
            __check((Files.readString(d1)).toString(), "data")
            __check((Files.readString(d2)).toString(), "data")
            Files.delete(src)
            Files.delete(d1)
            Files.delete(d2)
        }
