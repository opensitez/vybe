// vybe-test: kotlin/try_catch_flow/test_try_catch_in_loop
// origin: languages/kotlin/tests/kotlin/test_try_catch_flow.rs

fun maybe(i: Int): Int {
            if (i < 0) throw Exception("neg")
            return i
        }
        fun main() {
            var sum = 0
            for (i in -1..2) {
                try {
                    sum += maybe(i)
                } catch (e: Exception) {
                    sum += 10
                }
            }
            println(sum)
        }

