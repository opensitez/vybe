use super::helpers::{run_csharp, run_csharp_one};

#[test]
fn if_else() {
    let out = run_csharp(r#"
        var x = 10;
        if (x > 5) {
            Console.WriteLine("big");
        } else {
            Console.WriteLine("small");
        }
    "#);
    assert_eq!(out, vec!["big"]);
}

#[test]
fn if_elseif_else() {
    let out = run_csharp(r#"
        var x = 15;
        if (x > 20) { Console.WriteLine("big"); }
        else if (x > 10) { Console.WriteLine("medium"); }
        else { Console.WriteLine("small"); }
    "#);
    assert_eq!(out, vec!["medium"]);
}

#[test]
fn if_elseif_chain_multiple() {
    let out = run_csharp(r#"
        var x = 2;
        if (x == 1) { Console.WriteLine("one"); }
        else if (x == 2) { Console.WriteLine("two"); }
        else if (x == 3) { Console.WriteLine("three"); }
        else { Console.WriteLine("other"); }
    "#);
    assert_eq!(out, vec!["two"]);
}

#[test]
fn for_loop() {
    let out = run_csharp(r#"
        var sum = 0;
        for (var i = 1; i <= 5; i++) {
            sum = sum + i;
        }
        Console.WriteLine(sum);
    "#);
    assert_eq!(out, vec!["15"]);
}

#[test]
fn while_loop() {
    let out = run_csharp(r#"
        var i = 0;
        while (i < 3) {
            i = i + 1;
        }
        Console.WriteLine(i);
    "#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn do_while_loop() {
    let out = run_csharp(r#"
        var i = 0;
        do {
            i = i + 1;
        } while (i < 5);
        Console.WriteLine(i);
    "#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn foreach_array() {
    let out = run_csharp(r#"
        var sum = 0;
        foreach (var x in new int[] { 10, 20, 30 }) {
            sum = sum + x;
        }
        Console.WriteLine(sum);
    "#);
    assert_eq!(out, vec!["60"]);
}

#[test]
fn switch_basic() {
    let out = run_csharp(r#"
        var x = 2;
        var result = "";
        switch (x) {
            case 1: result = "one"; break;
            case 2: result = "two"; break;
            case 3: result = "three"; break;
            default: result = "other"; break;
        }
        Console.WriteLine(result);
    "#);
    assert_eq!(out, vec!["two"]);
}

#[test]
fn nested_for_loops() {
    let out = run_csharp(r#"
        var sum = 0;
        for (var i = 0; i < 3; i++) {
            for (var j = 0; j < 3; j++) {
                sum = sum + 1;
            }
        }
        Console.WriteLine(sum);
    "#);
    assert_eq!(out, vec!["9"]);
}

#[test]
fn break_in_loop() {
    let out = run_csharp(r#"
        var result = 0;
        for (var i = 0; i < 100; i++) {
            if (i == 5) break;
            result = result + 1;
        }
        Console.WriteLine(result);
    "#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn continue_in_loop() {
    let out = run_csharp(r#"
        var sum = 0;
        for (var i = 0; i < 10; i++) {
            if (i % 2 != 0) continue;
            sum = sum + i;
        }
        Console.WriteLine(sum);
    "#);
    assert_eq!(out, vec!["20"]);
}

#[test]
fn try_catch_basic() {
    let out = run_csharp(r#"
        try {
            throw new Exception("oops");
        } catch (Exception e) {
            Console.WriteLine("caught");
        }
    "#);
    assert_eq!(out, vec!["caught"]);
}
