// vybe-test: kotlin/kotlin_path_and_files/test_files_temp_write_read_delete_roundtrip
// origin: languages/kotlin/tests/kotlin/test_kotlin_path_and_files.rs

import java.nio.file.Files
        import java.nio.file.Path

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val tmp = Files.createTempFile("vybe", ".txt")
            Files.writeString(tmp, "hello")
            val text = Files.readString(tmp)
            __check((text).toString(), "hello")
            val moved = Files.move(tmp, tmp.resolveSibling("vybe_moved_" + tmp.fileName.toString()), java.nio.file.StandardCopyOption.REPLACE_EXISTING)
            __check((Files.exists(tmp)).toString(), "false")
            __check((Files.exists(moved)).toString(), "true")
            __check((Files.readString(moved)).toString(), "hello")
            Files.delete(moved)
        }
