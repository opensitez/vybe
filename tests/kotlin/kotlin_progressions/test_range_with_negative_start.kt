// vybe-test: kotlin/kotlin_progressions/test_range_with_negative_start
// origin: languages/kotlin/tests/kotlin/test_kotlin_progressions.rs

fun main() {
            var out = ""
            for (v in -1..2) {
                out = out + v.toString()
            }
            println(out)
        }

