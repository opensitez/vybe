// vybe-test: kotlin/local_returns/test_local_return_in_each_indexed
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun main() {
            val out = StringBuilder()
            listOf(1, 2, 3).forEachIndexed { index, value ->
                if (index == 1) return@forEachIndexed
                out.append(value)
            }
            println(out.toString())
        }

