// vybe-test: kotlin/local_functions/test_local_function_with_varargs_input
// origin: languages/kotlin/tests/kotlin/test_local_functions.rs

fun main() {
            fun join(prefix: String, vararg parts: Int): String {
                var out = prefix
                for (part in parts) {
                    out += ":" + part.toString()
                }
                return out
            }
            println(join("a", 1, 2, 3))
        }

