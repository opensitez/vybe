// vybe-test: kotlin/type_cast_edges/test_lambda_is_check_in_conditional
// origin: languages/kotlin/tests/kotlin/test_type_cast_edges.rs

fun main() {
            val value: Any = { s: String -> s.uppercase() }
            if (value is (String) -> String) {
                println(value("ab"))
            } else {
                println("no")
            }
        }

