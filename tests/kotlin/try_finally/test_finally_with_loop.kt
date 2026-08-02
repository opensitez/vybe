// vybe-test: kotlin/try_finally/test_finally_with_loop
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun main() {
        var x = 0
        for (i in 0..1) {
            try {
                x += i
            } finally {
                x += 10
            }
        }
        println(x)
    }

