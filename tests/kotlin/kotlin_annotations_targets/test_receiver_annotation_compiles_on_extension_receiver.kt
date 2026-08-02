// vybe-test: kotlin/kotlin_annotations_targets/test_receiver_annotation_compiles_on_extension_receiver
// origin: languages/kotlin/tests/kotlin/test_kotlin_annotations_targets.rs

@Target(AnnotationTarget.RECEIVER)
        annotation class ReceiverMarker

        class Box {
            fun text() = "ok"
        }

        @ReceiverMarker
        fun Box.announce(): String = this.text()

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Box().announce()).toString(), "ok")
        }
