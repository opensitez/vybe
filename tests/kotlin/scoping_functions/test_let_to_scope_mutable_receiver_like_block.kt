// vybe-test: kotlin/scoping_functions/test_let_to_scope_mutable_receiver_like_block
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

fun main() {
            val values = mutableListOf(1, 2, 3)
            val doubled = values.let {
                val out = mutableListOf<Int>()
                for (value in it) {
                    out.add(value * 2)
                }
                out
            }
            println(doubled.joinToString(","))
        }

