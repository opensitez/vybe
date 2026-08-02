// vybe-test: kotlin/kotlin_path_and_files/test_files_copy_and_delete_if_exists
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
            val src = Files.createTempFile("vybe-copy-src", ".txt")
            Files.writeString(src, "alpha")
            val dst = Files.createTempFile("vybe-copy-dst", ".txt")
            Files.copy(src, dst, StandardCopyOption.REPLACE_EXISTING)
            val before = Files.readString(dst)
            val removed = Files.deleteIfExists(src)
            __check((before).toString(), "alpha")
            __check((removed).toString(), "true")
            __check((Files.exists(src)).toString(), "false")
            Files.delete(dst)
        }
