// vybe-test: kotlin/labeled_control_flow/test_label_skips_only_after_conditions
// origin: languages/kotlin/tests/kotlin/test_labeled_control_flow.rs

fun main() {
            var out = ""
            outer@ for (ch in listOf('a', 'b', 'c')) {
                for (digit in listOf('1', '2', '3')) {
                    if (ch == 'b' && digit == '2') {
                        continue@outer
                    }
                    out += "$ch$digit|"
                }
            }
            println(out)
        }

