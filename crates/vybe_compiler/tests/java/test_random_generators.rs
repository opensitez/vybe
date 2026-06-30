use crate::helpers::run_main;

#[test]
fn random_next_int_bounded() {
    let out = run_main(
        "java.util.Random r = new java.util.Random(1); int v = r.nextInt(5); System.out.println(v >= 0 && v < 5);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn random_next_int_fixed_seed() {
    let out = run_main(
        "java.util.Random r = new java.util.Random(42); System.out.println(r.nextInt());",
    );
    assert_eq!(out, vec!["-1170105035"]);
}

#[test]
fn random_next_double_range() {
    let out = run_main(
        "java.util.Random r = new java.util.Random(0); double d = r.nextDouble(); System.out.println(d >= 0.0 && d < 1.0);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn random_next_boolean_seed() {
    let out = run_main(
        "java.util.Random r = new java.util.Random(123); System.out.println(r.nextBoolean());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn random_next_long_seed() {
    let out = run_main(
        "java.util.Random r = new java.util.Random(7); System.out.println(r.nextLong() != 0);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn random_next_float_seed() {
    let out = run_main(
        "java.util.Random r = new java.util.Random(9); System.out.println(r.nextFloat() >= 0.0f);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn random_next_bytes_length() {
    let out = run_main(
        "java.util.Random r = new java.util.Random(2); byte[] b = new byte[4]; r.nextBytes(b); System.out.println(b.length);",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn random_doubles_stream_count() {
    let out = run_main(
        "java.util.Random r = new java.util.Random(5); long c = r.doubles().limit(3).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn random_ints_stream_bounded() {
    let out = run_main(
        "java.util.Random r = new java.util.Random(6); long c = r.ints(0, 10).limit(5).filter(n -> n >= 0).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn random_longs_stream_limit() {
    let out = run_main(
        "java.util.Random r = new java.util.Random(8); long c = r.longs().limit(2).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn splittable_random_next_int() {
    let out = run_main(
        "java.util.SplittableRandom r = new java.util.SplittableRandom(11); System.out.println(r.nextInt(100) < 100);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn splittable_random_split_independent() {
    let out = run_main(
        "java.util.SplittableRandom parent = new java.util.SplittableRandom(12); java.util.SplittableRandom child = parent.split(); System.out.println(child.nextInt(5) >= 0);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn splittable_random_next_long() {
    let out = run_main(
        "java.util.SplittableRandom r = new java.util.SplittableRandom(13); System.out.println(r.nextLong() != 0L);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn splittable_random_next_double() {
    let out = run_main(
        "java.util.SplittableRandom r = new java.util.SplittableRandom(14); System.out.println(r.nextDouble() >= 0.0);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn splittable_random_ints_stream() {
    let out = run_main(
        "java.util.SplittableRandom r = new java.util.SplittableRandom(15); long c = r.ints(2).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn splittable_random_longs_stream() {
    let out = run_main(
        "java.util.SplittableRandom r = new java.util.SplittableRandom(16); long c = r.longs(3).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn splittable_random_doubles_stream() {
    let out = run_main(
        "java.util.SplittableRandom r = new java.util.SplittableRandom(17); long c = r.doubles(1).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn thread_local_random_current() {
    let out = run_main(
        "System.out.println(java.util.concurrent.ThreadLocalRandom.current().nextInt(10) < 10);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn thread_local_random_next_long() {
    let out = run_main(
        "System.out.println(java.util.concurrent.ThreadLocalRandom.current().nextLong() != 0L);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn thread_local_random_next_double() {
    let out = run_main(
        "System.out.println(java.util.concurrent.ThreadLocalRandom.current().nextDouble() >= 0.0);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn thread_local_random_bounded_int() {
    let out = run_main(
        "int v = java.util.concurrent.ThreadLocalRandom.current().nextInt(1, 4); System.out.println(v >= 1 && v < 4);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn random_set_seed_changes_sequence() {
    let out = run_main(
        "java.util.Random r = new java.util.Random(); r.setSeed(99); int a = r.nextInt(); r.setSeed(99); int b = r.nextInt(); System.out.println(a == b);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn splittable_random_next_int_origin_bound() {
    let out = run_main(
        "java.util.SplittableRandom r = new java.util.SplittableRandom(18); int v = r.nextInt(2, 5); System.out.println(v >= 2 && v < 5);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn random_ints_fixed_seed_sum_positive() {
    let out = run_main(
        "java.util.Random r = new java.util.Random(3); int s = r.ints(3, 1, 4).sum(); System.out.println(s > 0);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn splittable_random_doubles_limited_average() {
    let out = run_main(
        "java.util.SplittableRandom r = new java.util.SplittableRandom(19); double a = r.doubles(10).average().getAsDouble(); System.out.println(a >= 0.0 && a < 1.0);",
    );
    assert_eq!(out, vec!["true"]);
}

