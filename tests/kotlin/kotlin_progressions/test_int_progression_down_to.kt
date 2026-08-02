// vybe-test: kotlin/kotlin_progressions/test_int_progression_down_to
// origin: languages/kotlin/tests/kotlin/test_kotlin_progressions.rs

fun main() {
            var out = ""
            for (v in 5 downTo 1) {
                out = out + v.toString()
            }
            println(out)
        }

