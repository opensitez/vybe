// vybe-test: kotlin/local_returns/test_local_return_from_for_each
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun main() {
            val out = mutableListOf<Int>()
            listOf(1, 2, 3, 4).forEach {
                if (it == 3) return@forEach
                out.add(it)
            }
            println(out.joinToString(","))
        }

