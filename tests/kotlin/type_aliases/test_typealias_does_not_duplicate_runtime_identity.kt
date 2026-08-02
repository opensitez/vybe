// vybe-test: kotlin/type_aliases/test_typealias_does_not_duplicate_runtime_identity
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias Label = String

        class Box(var value: Label)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val first = Box("a")
            val second: Label = "a"
            first.value = second
            __check((first.value).toString(), "a")
        }
