// vybe-test: kotlin/kotlin_progressions/test_range_while_loop_equiv
// origin: languages/kotlin/tests/kotlin/test_kotlin_progressions.rs

fun main() {
            var i = 3
            val r = 1..3
            var out = ""
            while (i >= 1) {
                out = out + i.toString()
                i -= 1
            }
            println(r.toList().joinToString(","))
            println(out)
        }

