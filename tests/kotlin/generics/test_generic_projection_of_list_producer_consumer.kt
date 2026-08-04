// vybe-test: kotlin/generics/test_generic_projection_of_list_producer_consumer
// origin: languages/kotlin/tests/kotlin/test_generics.rs

interface Producer<out T> {
            fun produce(): T
        }

        interface Consumer<in T> {
            fun consume(value: T)
        }

        class StringProducer : Producer<String> {
            var value = "go"
            override fun produce(): String = value
        }

        class AnyConsumer : Consumer<Any> {
            var last: Any? = null
            override fun consume(value: Any) { last = value }
            fun seen(): String = last.toString()
        }

        fun pipe(source: Producer<String>, sink: Consumer<CharSequence>) {
            sink.consume(source.produce())
        }

        var __buf: String = ""

fun __p(s: String) {
    __buf = __buf + s + "\n"
}

fun __pr(s: String) {
    __buf = __buf + s
}

// The final `println` contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted. Written as two equality
// tests rather than trimming: `String.endsWith` is not implemented in Vybe's
// Kotlin (measured — `"ab\n".endsWith("\n")` throws "undefined is not
// callable"), and a harness that cannot run asserts nothing at all. The cargo
// helper split on "\n" and popped trailing empties, so the two forms were
// equivalent there too.
fun __check(want: String) {
    if (__buf != want && __buf != want + "\n") {
        println("FAIL: want [" + want + "] got [" + __buf + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val producer: Producer<String> = StringProducer()
            val consumer = AnyConsumer()
            pipe(producer, consumer)
            __p((consumer.seen()).toString())
        
__check("go")
}
