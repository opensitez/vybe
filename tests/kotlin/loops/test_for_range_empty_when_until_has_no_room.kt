// vybe-test: kotlin/loops/test_for_range_empty_when_until_has_no_room
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun main() {
            var sum = 0
            for (i in 3 until 3) {
                sum += i
            }
            println(sum)
        }

