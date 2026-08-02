// vybe-test: kotlin/kotlin_progressions/test_int_progression_step_down_to
// origin: languages/kotlin/tests/kotlin/test_kotlin_progressions.rs

fun main() {
            val values = 10 downTo 1 step 3
            var out = ""
            for (v in values) {
                out = out + v.toString()
                if (v > 1) out = out + ","
            }
            println(out)
        }

