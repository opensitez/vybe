// vybe-test: kotlin/kotlin_char_range_ops/test_descending_char_range
// origin: languages/kotlin/tests/kotlin/test_kotlin_char_range_ops.rs

fun main() {
            var out = ""
            for (c in 'e' downTo 'c') {
                out += c
            }
            println(out)
        }

