// vybe-test: kotlin/kotlin_progressions/test_int_progression_until
// origin: languages/kotlin/tests/kotlin/test_kotlin_progressions.rs

fun main() {
            val values = 1 until 4
            var out = ""
            for (v in values) { out = out + v.toString() }
            println(out)
            println(values.last)
        }

