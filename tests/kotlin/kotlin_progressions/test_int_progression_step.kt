// vybe-test: kotlin/kotlin_progressions/test_int_progression_step
// origin: languages/kotlin/tests/kotlin/test_kotlin_progressions.rs

fun main() {
            val values = 1..10 step 2
            var out = ""
            for (v in values) {
                out = out + v.toString()
                out = out + ","
            }
            println(out)
        }

