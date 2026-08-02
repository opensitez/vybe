// vybe-test: kotlin/type_aliases/test_typealias_aliases_with_nullable_generic_type
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias Maybe<T> = T?
        typealias MaybeText = Maybe<String>

        fun fallback(value: MaybeText, default: String): String {
            return value ?: default
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val present: MaybeText = "ok"
            val missing: MaybeText = null
            __check((fallback(present, "none")).toString(), "ok")
            __check((fallback(missing, "none")).toString(), "none")
        }
