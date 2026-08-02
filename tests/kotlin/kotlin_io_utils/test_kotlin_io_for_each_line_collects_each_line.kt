// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_for_each_line_collects_each_line
// origin: languages/kotlin/tests/kotlin/test_kotlin_io_utils.rs

fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_foreach_" + System.nanoTime() + ".txt")
            file.writeText("u\nv\nw")
            val joined = StringBuilder()
            file.forEachLine { joined.append(it).append(".") }
            println(joined.toString())
            file.delete()
        }

