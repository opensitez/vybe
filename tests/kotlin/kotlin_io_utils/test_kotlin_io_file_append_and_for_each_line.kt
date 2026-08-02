// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_file_append_and_for_each_line
// origin: languages/kotlin/tests/kotlin/test_kotlin_io_utils.rs

fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_append_lines_" + System.nanoTime() + ".txt")
            file.writeText("1\n")
            file.appendText("2\n")
            file.appendText("3\n")
            val total = file.readText().trim().split("\n").size
            val first = StringBuilder()
            file.forEachLine { first.append(it) }
            println(total)
            println(first.toString())
            file.delete()
        }

