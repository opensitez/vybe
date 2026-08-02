// vybe-test: kotlin/local_returns/test_local_return_from_while_like_loop
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun main() {
            var i = 0
            val out = StringBuilder()
            while (i < 5) {
                if (i == 3) { i += 1
continue }
                out.append(i)
                i += 1
            }
            println(out.toString())
        }

