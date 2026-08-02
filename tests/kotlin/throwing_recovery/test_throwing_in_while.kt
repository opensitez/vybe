// vybe-test: kotlin/throwing_recovery/test_throwing_in_while
// origin: languages/kotlin/tests/kotlin/test_throwing_recovery.rs

fun main() {
            var i = 0
            while (i < 3) {
                try {
                    if (i == 1) throw Exception("x")
                    println(i)
                } catch (e: Exception) {
                    println("err")
                }
                i += 1
            }
        }

