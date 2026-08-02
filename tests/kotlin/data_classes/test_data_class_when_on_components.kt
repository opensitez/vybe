// vybe-test: kotlin/data_classes/test_data_class_when_on_components
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Kind(val id: Int, val name: String)

        fun classify(kind: Kind): String {
            val (id, name) = kind
            return if (id == 1) "first:" + name else "other:" + name
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val first = Kind(1, "root")
            val second = Kind(2, "leaf")
            __check((classify(first)).toString(), "first:root")
            __check((classify(second)).toString(), "other:leaf")
        }
