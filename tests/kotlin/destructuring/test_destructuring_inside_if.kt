// vybe-test: kotlin/destructuring/test_destructuring_inside_if
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun main() {
            if (true) {
                val (a, b) = Pair(7, 8)
                println(a)
                println(b)
            } else {
                println("no")
            }
        }

