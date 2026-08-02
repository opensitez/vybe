// vybe-test: kotlin/kotlin_path_and_files/test_directory_walk_with_filter
// origin: languages/kotlin/tests/kotlin/test_kotlin_path_and_files.rs

import java.nio.file.Files
        import java.nio.file.Paths

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = Paths.get(java.lang.System.getProperty("java.io.tmpdir"), "vybe_walk_" + System.nanoTime().toString())
            val a = Files.createDirectories(base.resolve("a"))
            val b = base.resolve("a.txt")
            val c = base.resolve("b.log")
            Files.writeString(b, "one")
            Files.writeString(c, "two")
            val count = Files.list(base).filter { p -> Files.isRegularFile(p) }.count().toInt()
            __check((count).toString(), "2")
            val hasTxt = Files.newDirectoryStream(base, "*.txt").use {
                it.asSequence().count()
            }
            __check((hasTxt).toString(), "1")
            Files.delete(b)
            Files.delete(c)
            Files.delete(a)
            Files.delete(base)
        }
