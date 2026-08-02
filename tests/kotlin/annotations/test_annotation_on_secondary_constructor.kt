// vybe-test: kotlin/annotations/test_annotation_on_secondary_constructor
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

class Session {
    val id: Int

    constructor() {
        this.id = 1
    }

    @Deprecated("secondary")
    constructor(id: Int) {
        this.id = id
    }
}

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
    val s = Session(5)
    __check((s.id).toString(), "5")
}
