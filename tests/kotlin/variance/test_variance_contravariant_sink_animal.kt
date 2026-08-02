// vybe-test: kotlin/variance/test_variance_contravariant_sink_animal
// origin: languages/kotlin/tests/kotlin/test_variance.rs

interface Sink<in T> { fun consume(value: T) }
        open class Animal { fun label() = "a" }
        class Recorder : Sink<Animal> {
            override fun consume(value: Animal) { __check((value.label()).toString(), "d") }
        }
        class Dog : Animal() { override fun label() = "d" }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val sink: Sink<Dog> = Recorder()
            sink.consume(Dog())
        }
