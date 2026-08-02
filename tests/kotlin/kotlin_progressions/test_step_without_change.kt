// vybe-test: kotlin/kotlin_progressions/test_step_without_change
// origin: languages/kotlin/tests/kotlin/test_kotlin_progressions.rs

fun main() {
            val values = 1..10 step 1
            var x = 0
            for (v in values) { x = v }
            println(x)
            val empty = 10 downTo 12 step 2
            println(empty.toList().size)
        }

