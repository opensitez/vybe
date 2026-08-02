// vybe-test: kotlin/class_delegation/test_collection_delegation_with_generic_type
// origin: languages/kotlin/tests/kotlin/test_class_delegation.rs

interface Labeler<T> { fun label(value: T): String }

        class StringLabeler : Labeler<String> {
            override fun label(value: String) = "[$value]"
        }

        class DelegatingLabeler(delegate: Labeler<String>) : Labeler<String> by delegate

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val l = DelegatingLabeler(StringLabeler())
            __check((l.label("x")).toString(), "[x]")
        }
