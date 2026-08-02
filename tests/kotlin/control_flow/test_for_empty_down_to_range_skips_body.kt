// vybe-test: kotlin/control_flow/test_for_empty_down_to_range_skips_body
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            var seen = ""
            for (i in 5 downTo 10) {
                seen += i.toString()
            }
            println(seen.isEmpty())
        }

