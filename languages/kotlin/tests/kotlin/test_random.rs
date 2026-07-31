use crate::helpers::run_prints;

#[test]
fn test_random_default_generator_exists() {
    let out = run_prints(r#"
        fun main() {
            println(kotlin.random.Random.Default != null)
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_random_seeded_int_sequence_is_repeatable() {
    let out = run_prints(r#"
        fun main() {
            val a = kotlin.random.Random(7)
            val b = kotlin.random.Random(7)
            println(a.nextInt() == b.nextInt())
            println(a.nextInt() == b.nextInt())
        }
    "#);
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_random_seeded_long_sequence_is_repeatable() {
    let out = run_prints(r#"
        fun main() {
            val a = kotlin.random.Random(11)
            val b = kotlin.random.Random(11)
            println(a.nextLong() == b.nextLong())
            println(a.nextLong() == b.nextLong())
        }
    "#);
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_random_seeded_double_sequence_is_repeatable() {
    let out = run_prints(r#"
        fun main() {
            val a = kotlin.random.Random(13)
            val b = kotlin.random.Random(13)
            println(a.nextDouble() == b.nextDouble())
            println(a.nextDouble() == b.nextDouble())
        }
    "#);
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_random_seeded_float_sequence_is_repeatable() {
    let out = run_prints(r#"
        fun main() {
            val a = kotlin.random.Random(17)
            val b = kotlin.random.Random(17)
            println(a.nextFloat() == b.nextFloat())
            println(a.nextFloat() == b.nextFloat())
        }
    "#);
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_random_seeded_boolean_sequence_is_repeatable() {
    let out = run_prints(r#"
        fun main() {
            val a = kotlin.random.Random(19)
            val b = kotlin.random.Random(19)
            println(a.nextBoolean() == b.nextBoolean())
            println(a.nextBoolean() == b.nextBoolean())
            println(a.nextBoolean() == b.nextBoolean())
        }
    "#);
    assert_eq!(out, &["true", "true", "true"]);
}

#[test]
fn test_random_next_int_with_upper_bound_uses_bound() {
    let out = run_prints(r#"
        fun main() {
            val r = kotlin.random.Random(3)
            println(r.nextInt(10) >= 0)
            println(r.nextInt(10) < 10)
            println(r.nextInt(10) >= 0 && r.nextInt(10) < 10)
        }
    "#);
    assert_eq!(out, &["true", "true", "true"]);
}

#[test]
fn test_random_next_int_with_lower_and_upper_bound() {
    let out = run_prints(r#"
        fun main() {
            val r = kotlin.random.Random(5)
            val first = r.nextInt(-4, 7)
            val second = r.nextInt(-4, 7)
            val third = r.nextInt(-4, 7)
            println(first in -4..6)
            println(second in -4..6)
            println(third in -4..6)
        }
    "#);
    assert_eq!(out, &["true", "true", "true"]);
}

#[test]
fn test_random_next_long_with_bounds_stays_in_range() {
    let out = run_prints(r#"
        fun main() {
            val r = kotlin.random.Random(23)
            val value = r.nextLong(-10L, 20L)
            println(value >= -10)
            println(value < 20)
            println(value in -10L..19L)
        }
    "#);
    assert_eq!(out, &["true", "true", "true"]);
}

#[test]
fn test_random_next_long_with_upper_bound() {
    let out = run_prints(r#"
        fun main() {
            val r = kotlin.random.Random(29)
            val v = r.nextLong(100L)
            println(v >= 0)
            println(v < 100)
            println(v in 0L..99L)
        }
    "#);
    assert_eq!(out, &["true", "true", "true"]);
}

#[test]
fn test_random_next_double_default_is_unit_interval() {
    let out = run_prints(r#"
        fun main() {
            val r = kotlin.random.Random(31)
            val d = r.nextDouble()
            println(d >= 0.0)
            println(d < 1.0)
        }
    "#);
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_random_next_double_with_upper_bound() {
    let out = run_prints(r#"
        fun main() {
            val r = kotlin.random.Random(37)
            val d = r.nextDouble(1.5)
            println(d >= 0.0)
            println(d < 1.5)
        }
    "#);
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_random_next_double_with_bounds() {
    let out = run_prints(r#"
        fun main() {
            val r = kotlin.random.Random(41)
            val d = r.nextDouble(-1.0, 2.0)
            println(d >= -1.0)
            println(d < 2.0)
        }
    "#);
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_random_next_float_default_in_unit_interval() {
    let out = run_prints(r#"
        fun main() {
            val r = kotlin.random.Random(43)
            val f = r.nextFloat()
            println(f >= 0.0f)
            println(f < 1.0f)
        }
    "#);
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_random_next_bytes_has_expected_length() {
    let out = run_prints(r#"
        fun main() {
            val r = kotlin.random.Random(47)
            println(r.nextBytes(3).size)
            println(r.nextBytes(5).size)
        }
    "#);
    assert_eq!(out, &["3", "5"]);
}

#[test]
fn test_random_next_bytes_repeatable_with_seed() {
    let out = run_prints(r#"
        fun main() {
            val a = kotlin.random.Random(53)
            val b = kotlin.random.Random(53)
            val ba = a.nextBytes(4)
            val bb = b.nextBytes(4)
            val same = ba.joinToString(",") == bb.joinToString(",")
            println(same)
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_random_next_bytes_content_varies_after_consumption() {
    let out = run_prints(r#"
        fun main() {
            val r = kotlin.random.Random(59)
            val first = r.nextBytes(2)
            val second = r.nextBytes(2)
            println(first.joinToString(",") == second.joinToString(","))
        }
    "#);
    assert_eq!(out, &["false"]);
}

#[test]
fn test_random_list_shuffle_with_seed_is_repeatable() {
    let out = run_prints(r#"
        fun main() {
            val a = kotlin.random.Random(61)
            val b = kotlin.random.Random(61)
            val sourceA = mutableListOf(1, 2, 3, 4, 5)
            val sourceB = mutableListOf(1, 2, 3, 4, 5)
            sourceA.shuffle(a)
            sourceB.shuffle(b)
            println(sourceA == sourceB)
            println(sourceA.size)
        }
    "#);
    assert_eq!(out, &["true", "5"]);
}

#[test]
fn test_random_list_shuffled_with_seed_is_repeatable() {
    let out = run_prints(r#"
        fun main() {
            val a = kotlin.random.Random(67).shuffled(listOf(1, 2, 3, 4))
            val b = kotlin.random.Random(67).shuffled(listOf(1, 2, 3, 4))
            println(a == b)
            println(a.size)
        }
    "#);
    assert_eq!(out, &["true", "4"]);
}

#[test]
fn test_random_choose_with_reproducible_seed() {
    let out = run_prints(r#"
        fun main() {
            val src = listOf("a", "b", "c", "d")
            val a = src.random(kotlin.random.Random(71))
            val b = src.random(kotlin.random.Random(71))
            println(a == b)
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_random_choice_from_range_is_in_bounds() {
    let out = run_prints(r#"
        fun main() {
            val picked = kotlin.random.Random(73).nextInt(2..9)
            println(picked in 2..9)
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_random_next_int_exclusive_bound_rejects_invalid() {
    let out = run_prints(r#"
        fun main() {
            try {
                kotlin.random.Random(79).nextInt(0)
                println("ok")
            } catch (ex: IllegalArgumentException) {
                println("error")
            }
        }
    "#);
    assert_eq!(out, &["error"]);
}

#[test]
fn test_random_next_int_range_rejects_invalid_order() {
    let out = run_prints(r#"
        fun main() {
            try {
                kotlin.random.Random(83).nextInt(10, 2)
                println("ok")
            } catch (ex: IllegalArgumentException) {
                println("error")
            }
        }
    "#);
    assert_eq!(out, &["error"]);
}

#[test]
fn test_random_next_long_range_rejects_invalid_order() {
    let out = run_prints(r#"
        fun main() {
            try {
                kotlin.random.Random(89).nextLong(3L, -3L)
                println("ok")
            } catch (ex: IllegalArgumentException) {
                println("error")
            }
        }
    "#);
    assert_eq!(out, &["error"]);
}

#[test]
fn test_random_next_double_range_rejects_invalid_order() {
    let out = run_prints(r#"
        fun main() {
            try {
                kotlin.random.Random(97).nextDouble(2.0, 1.0)
                println("ok")
            } catch (ex: IllegalArgumentException) {
                println("error")
            }
        }
    "#);
    assert_eq!(out, &["error"]);
}

#[test]
fn test_random_repeatability_with_default_seeded_factory() {
    let out = run_prints(r#"
        fun main() {
            val a = kotlin.random.Random(101)
            val b = kotlin.random.Random(101)
            println(a.nextInt(0, 1000) == b.nextInt(0, 1000))
            println(a.nextLong(0L, 1000L) == b.nextLong(0L, 1000L))
            println(a.nextBoolean() == b.nextBoolean())
        }
    "#);
    assert_eq!(out, &["true", "true", "true"]);
}

#[test]
fn test_random_moving_state_advances_sequence() {
    let out = run_prints(r#"
        fun main() {
            val r = kotlin.random.Random(103)
            val first = r.nextInt()
            val second = r.nextInt()
            val third = kotlin.random.Random(103).nextInt()
            println(first == second)
            println(first == third)
        }
    "#);
    assert_eq!(out, &["false", "true"]);
}

#[test]
fn test_random_bool_runs_across_many_calls() {
    let out = run_prints(r#"
        fun main() {
            val r1 = kotlin.random.Random(107)
            val r2 = kotlin.random.Random(107)
            var trueCount = 0
            var i = 0
            while (i < 8) {
                if (r1.nextBoolean()) trueCount++
                i++
            }
            var trueCount2 = 0
            var j = 0
            while (j < 8) {
                if (r2.nextBoolean()) trueCount2++
                j++
            }
            println(trueCount == trueCount2)
            println(trueCount in 0..8)
        }
    "#);
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_random_double_has_fractional_precision() {
    let out = run_prints(r#"
        fun main() {
            val r = kotlin.random.Random(109)
            val value = r.nextDouble()
            println(value.toString().contains("."))
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_random_int_and_double_repeatability_for_same_seed() {
    let out = run_prints(r#"
        fun main() {
            val a = kotlin.random.Random(113)
            val b = kotlin.random.Random(113)
            val aFirst = a.nextInt(1000)
            val aSecond = a.nextInt(1000)
            val bFirst = b.nextInt(1000)
            val bSecond = b.nextInt(1000)
            val aDouble = a.nextDouble()
            val bDouble = b.nextDouble()
            println(aFirst == bFirst)
            println(aSecond == bSecond)
            println(aDouble == bDouble)
        }
    "#);
    assert_eq!(out, &["true", "true", "true"]);
}
