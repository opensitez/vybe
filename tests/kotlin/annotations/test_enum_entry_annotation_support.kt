// vybe-test: kotlin/annotations/test_enum_entry_annotation_support
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

annotation class State

        enum class Mode {
            @State
            OFF,

            @State
            ON
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Mode.ON.name).toString(), "ON")
        }
