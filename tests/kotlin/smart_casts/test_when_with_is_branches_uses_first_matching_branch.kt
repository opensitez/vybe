// vybe-test: kotlin/smart_casts/test_when_with_is_branches_uses_first_matching_branch
// origin: languages/kotlin/tests/kotlin/test_smart_casts.rs

interface Shape
        class Circle : Shape
        class Square : Shape

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Shape = Circle()
            val label = when (value) {
                is Circle -> "circle"
                is Square -> "square"
                else -> "other"
            }
            __check((label).toString(), "circle")
        }
