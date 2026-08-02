// vybe-test: kotlin/type_aliases/test_typealias_for_set_with_factory_function
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias NameSet = HashSet<String>

        fun make(): NameSet {
            return NameSet(listOf("x", "y", "x"))
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((make().size).toString(), "2")
        }
