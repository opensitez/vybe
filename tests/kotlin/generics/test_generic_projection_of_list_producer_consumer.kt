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

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val producer: Producer<String> = StringProducer()
            val consumer = AnyConsumer()
            pipe(producer, consumer)
            __check((consumer.seen()).toString(), "go")
        }
