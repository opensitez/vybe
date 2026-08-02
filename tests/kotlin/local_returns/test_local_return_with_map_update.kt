// vybe-test: kotlin/local_returns/test_local_return_with_map_update
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun main() {
            val out = mutableMapOf<Int, Int>()
            for (i in 1..4) {
                if (i == 3) continue
                out[i] = i
            }
            println(out.size)
        }

