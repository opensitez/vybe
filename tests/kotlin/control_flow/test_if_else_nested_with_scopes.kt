// vybe-test: kotlin/control_flow/test_if_else_nested_with_scopes
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            val score = 82
            if (score >= 70) {
                if (score >= 80) {
                    println("pass-a")
                } else {
                    println("pass-b")
                }
            } else {
                println("fail")
            }
        }

