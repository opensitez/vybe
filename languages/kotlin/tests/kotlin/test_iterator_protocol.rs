use crate::helpers::run_prints;

#[test]
fn test_list_iterator_next_and_has_next() {
    let out = run_prints(r#"
        fun main() {
            val it = listOf(1, 2, 3).iterator()
            val b1 = it.hasNext()
            val v1 = it.next()
            val b2 = it.hasNext()
            val v2 = it.next()
            println(b1)
            println(v1)
            println(b2)
            println(v2)
            println(it.next())
        }
    "#);
    assert_eq!(out, &["true", "1", "true", "2", "3"]);
}

#[test]
fn test_iterator_for_each_contract() {
    let out = run_prints(r#"
        fun main() {
            val it = listOf("a", "b").iterator()
            it.forEachRemaining { println(it) }
        }
    "#);
    assert_eq!(out, &["a", "b"]);
}

#[test]
fn test_iterable_for_each_index() {
    let out = run_prints(r#"
        fun main() {
            val values = listOf(1, 2, 3)
            val out = StringBuilder()
            for ((i, v) in values.withIndex()) {
                out.append(i).append(":").append(v).append("|")
            }
            println(out.toString())
        }
    "#);
    assert_eq!(out, &["0:1|1:2|2:3|"]);
}

#[test]
fn test_iterator_over_set_is_unique() {
    let out = run_prints(r#"
        fun main() {
            val seen = linkedSetOf(1, 2, 3)
            val out = StringBuilder()
            for (v in seen) {
                out.append(v)
            }
            println(out.toString())
        }
    "#);
    assert_eq!(out, &["123"]);
}

#[test]
fn test_iterator_over_map_keys() {
    let out = run_prints(r#"
        fun main() {
            val map = linkedMapOf("a" to 1, "b" to 2)
            val values = StringBuilder()
            for (k in map.keys) {
                values.append(k)
            }
            println(values.toString())
            println(map.keys.count())
        }
    "#);
    assert_eq!(out, &["ab", "2"]);
}

#[test]
fn test_iterator_over_map_entries() {
    let out = run_prints(r#"
        fun main() {
            val map = mapOf("a" to 1, "b" to 2)
            val values = map.entries.joinToString("|") { it.key + ":" + it.value }
            println(values)
            println(map.entries.size)
        }
    "#);
    assert_eq!(out, &["a:1|b:2", "2"]);
}

#[test]
fn test_iterator_over_string_chars() {
    let out = run_prints(r#"
        fun main() {
            val it = "ab".iterator()
            var first = it.next()
            var second = it.next()
            println(first)
            println(second)
            println(it.hasNext())
        }
    "#);
    assert_eq!(out, &["a", "b", "false"]);
}

#[test]
fn test_custom_iterator_implementing_interface() {
    let out = run_prints(r#"
        class RangeIterator : Iterator<Int> {
            private var i = 0
            private val end = 3
            override fun hasNext() = i < end
            override fun next(): Int {
                val value = i
                i += 1
                return value
            }
        }

        class RangeIterable : Iterable<Int> {
            override fun iterator(): Iterator<Int> = RangeIterator()
        }

        fun main() {
            val it = RangeIterable().iterator()
            var sum = 0
            for (value in RangeIterable()) {
                sum += value
            }
            println(sum)
            println(it.hasNext())
            println(it.next())
            println(it.next())
        }
    "#);
    assert_eq!(out, &["3", "true", "1", "2"]);
}

#[test]
fn test_iterator_peeking_behavior() {
    let out = run_prints(r#"
        class P : Iterator<Int> {
            private val data = listOf(7, 8)
            private var index = 0
            override fun hasNext() = index < data.size
            override fun next() = data[index++]
        }

        fun main() {
            val it = P()
            println(it.hasNext())
            println(it.next())
            println(it.hasNext())
            println(it.next())
            println(it.hasNext())
        }
    "#);
    assert_eq!(out, &["true", "7", "true", "8", "false"]);
}

#[test]
fn test_iterator_multiple_iterables_independent() {
    let out = run_prints(r#"
        fun main() {
            val list = listOf(1, 2)
            val a = list.iterator()
            val b = list.iterator()
            println(a.next())
            println(b.next())
            println(a.next())
            println(b.next())
        }
    "#);
    assert_eq!(out, &["1", "1", "2", "2"]);
}

#[test]
fn test_iterator_can_consume_generator_with_sequence() {
    let out = run_prints(r#"
        fun main() {
            val seq = generateSequence(1) { it + 1 }.take(3)
            val it = seq.iterator()
            println(it.next())
            println(it.next())
            println(it.next())
            println(it.hasNext())
        }
    "#);
    assert_eq!(out, &["1", "2", "3", "false"]);
}

#[test]
fn test_iterator_drop_while_like_manual() {
    let out = run_prints(r#"
        fun main() {
            val it = listOf(1, 2, 3, 4).iterator()
            while (it.hasNext()) {
                val n = it.next()
                if (n < 3) continue
                print(n)
            }
            println("")
        }
    "#);
    assert_eq!(out, &["34"]);
}

#[test]
fn test_iterator_throwing_when_empty_next() {
    let out = run_prints(r#"
        fun main() {
            val it = emptyList<Int>().iterator()
            try {
                it.next()
                println("no")
            } catch (e: NoSuchElementException) {
                println("error")
            }
        }
    "#);
    assert_eq!(out, &["error"]);
}

#[test]
fn test_iterator_collect_to_mutable_list() {
    let out = run_prints(r#"
        fun main() {
            val src = listOf(4, 5, 6).iterator()
            val dst = mutableListOf<Int>()
            while (src.hasNext()) {
                dst.add(src.next())
            }
            println(dst.joinToString(","))
        }
    "#);
    assert_eq!(out, &["4,5,6"]);
}

#[test]
fn test_iterator_yield_sum_reduce() {
    let out = run_prints(r#"
        fun main() {
            val it = (1..4).iterator()
            var total = 0
            while (it.hasNext()) {
                total += it.next()
            }
            println(total)
        }
    "#);
    assert_eq!(out, &["10"]);
}

#[test]
fn test_iterable_sequence_chain() {
    let out = run_prints(r#"
        fun main() {
            val it = generateSequence(1) { it + 2 }
                .take(4)
                .toList()
                .iterator()
            println(it.next())
            println(it.next())
            println(it.next())
            println(it.next())
            println(it.hasNext())
        }
    "#);
    assert_eq!(out, &["1", "3", "5", "7", "false"]);
}

#[test]
fn test_iterator_map_side_effect_order() {
    let out = run_prints(r##"
        fun main() {
            val source = mutableListOf(1, 2, 3)
            var log = ""
            val it = source.map {
                log += "#" + it
                it
            }.iterator()
            println(it.next())
            println(it.next())
            println(log)
        }
    "##);
    assert_eq!(out, &["1", "2", "#1#2#3"]);
}

#[test]
fn test_iterator_flat_map_concats() {
    let out = run_prints(r#"
        fun main() {
            val it = listOf(listOf(1, 2), listOf(3)).flatMap { it }.iterator()
            val values = mutableListOf<Int>()
            while (it.hasNext()) values.add(it.next())
            println(values.joinToString(","))
        }
    "#);
    assert_eq!(out, &["1,2,3"]);
}

#[test]
fn test_iterator_with_for_each() {
    let out = run_prints(r#"
        fun main() {
            var acc = 0
            for (v in listOf(1, 2, 3)) {
                acc += v
            }
            println(acc)
        }
    "#);
    assert_eq!(out, &["6"]);
}

#[test]
fn test_iterator_filter_then_to_set() {
    let out = run_prints(r#"
        fun main() {
            val filtered = (1..5).asSequence().iterator().asSequence().filter { it % 2 == 0 }.toSet()
            println(filtered.joinToString(","))
        }
    "#);
    assert_eq!(out, &["2,4"]);
}

#[test]
fn test_iterator_drop_while_to_list() {
    let out = run_prints(r#"
        fun main() {
            val values = generateSequence(0) { it + 1 }
                .take(6)
                .dropWhile { it < 3 }
                .toList()
            println(values.joinToString(","))
        }
    "#);
    assert_eq!(out, &["3,4,5"]);
}

#[test]
fn test_iterator_mutable_source_snapshot() {
    let out = run_prints(r#"
        fun main() {
            val src = mutableListOf(1, 2, 3)
            val it = src.iterator()
            println(it.next())
            src.add(4)
            println(it.next())
            println(it.next())
        }
    "#);
    assert_eq!(out, &["1", "2", "3"]);
}

#[test]
fn test_iterator_reusable_with_new_instance() {
    let out = run_prints(r#"
        fun main() {
            val src = listOf(1, 2)
            val a = src.iterator()
            val b = src.iterator()
            println(a.toList().joinToString(","))
            println(b.toList().joinToString(","))
        }
    "#);
    assert_eq!(out, &["1,2", "1,2"]);
}
