// vybe-test: kotlin/type_aliases/test_typealias_extension_function_on_aliased_set_type
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias NameSet = MutableSet<String>

        fun NameSet.sortedSignature(): String {
            return this.toList().sorted().joinToString("|")
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val names: NameSet = hashSetOf("z", "a", "m")
            names.add("c")
            __check((names.sortedSignature()).toString(), "a|c|m|z")
        }
