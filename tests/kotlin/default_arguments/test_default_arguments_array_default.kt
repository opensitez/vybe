// vybe-test: kotlin/default_arguments/test_default_arguments_array_default
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

fun values(head: Int, rest: IntArray = intArrayOf(1, 2)): Int {
            var sum = head
            for (v in rest) { sum += v }
            return sum
        }
        fun main() {
            println(values(1))
            println(values(1, intArrayOf(5)))
        }

