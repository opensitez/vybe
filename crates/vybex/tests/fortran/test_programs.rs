use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// Fortran: Complete programs — algorithms, patterns
// ═══════════════════════════════════════════════════════════

#[test]
fn fizzbuzz() {
    compile_ok(r#"
program fizzbuzz
    integer :: i
    do i = 1, 15
        if (mod(i, 15) == 0) then
            print *, "FizzBuzz"
        else if (mod(i, 3) == 0) then
            print *, "Fizz"
        else if (mod(i, 5) == 0) then
            print *, "Buzz"
        else
            print *, i
        end if
    end do
end program fizzbuzz
"#);
}

#[test]
fn sum_of_squares() {
    let out = run_prints(r#"
program test
    integer :: i, total
    total = 0
    do i = 1, 10
        total = total + i * i
    end do
    print *, total
end program test
"#);
    assert_eq!(out, vec!["385"]);
}

#[test]
fn fibonacci() {
    compile_ok(r#"
program fib
    integer :: n, a, b, temp, i
    n = 10
    a = 0
    b = 1
    do i = 1, n
        temp = a + b
        a = b
        b = temp
    end do
    print *, a
end program fib
"#);
}

#[test]
fn temperature_conversion() {
    let out = run_prints(r#"
program test
    real :: celsius, fahrenheit
    celsius = 100.0
    fahrenheit = celsius * 9.0 / 5.0 + 32.0
    print *, fahrenheit
end program test
"#);
    assert_eq!(out, vec!["212"]);
}

#[test]
fn quadratic_formula() {
    compile_ok(r#"
program quadratic
    real :: a, b, c, discriminant, x1, x2
    a = 1.0
    b = -5.0
    c = 6.0
    discriminant = b**2 - 4.0*a*c
    if (discriminant >= 0.0) then
        x1 = (-b + sqrt(discriminant)) / (2.0*a)
        x2 = (-b - sqrt(discriminant)) / (2.0*a)
        print *, x1
        print *, x2
    end if
end program quadratic
"#);
}

#[test]
fn simple_statistics() {
    compile_ok(r#"
program stats
    integer :: i, n
    real :: sum, mean
    n = 5
    sum = 0.0
    do i = 1, n
        sum = sum + real(i)
    end do
    mean = sum / real(n)
    print *, "Mean:", mean
end program stats
"#);
}

#[test]
fn power_function() {
    let out = run_prints(r#"
program test
    print *, 2 ** 8
end program test
"#);
    assert_eq!(out, vec!["256"]);
}

#[test]
fn countdown() {
    compile_ok(r#"
program countdown
    integer :: i
    do i = 10, 1, -1
        print *, i
    end do
    print *, "Launch!"
end program countdown
"#);
}

#[test]
fn accumulator_pattern() {
    let out = run_prints(r#"
program test
    integer :: i, sum, product
    sum = 0
    product = 1
    do i = 1, 5
        sum = sum + i
        product = product * i
    end do
    print *, sum
    print *, product
end program test
"#);
    assert_eq!(out, vec!["15", "120"]);
}
