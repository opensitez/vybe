// vybe-test: kotlin/loops/test_for_over_string_iterates_code_points
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun main() {
            var out = ""
            for (ch in "K1") {
                out += ch.uppercase()
            }
            println(out)
        }

