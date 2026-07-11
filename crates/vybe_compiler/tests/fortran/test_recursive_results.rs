use super::helpers::compile_ok;
macro_rules! c {
    ($n:ident,$s:expr) => {
        #[test]
        fn $n() {
            compile_ok($s);
        }
    };
}
c!(
    recursive_results_01,
    "recursive integer function f(n) result(r)
integer::n
if (n<=1) then
 r=1
else
 r=f(n-1)
end if
end function f
"
);
c!(
    recursive_results_02,
    "recursive real function f(x) result(r)
real::x
r=x
end function f
"
);
c!(
    recursive_results_03,
    "recursive complex function f(x) result(r)
complex::x
r=x
end function f
"
);
c!(
    recursive_results_04,
    "recursive character(len=3) function f() result(r)
r='abc'
end function f
"
);
c!(
    recursive_results_05,
    "recursive logical function f() result(r)
r=.true.
end function f
"
);
c!(
    recursive_results_06,
    "recursive integer function f(n) result(r)
integer::n
r=n
end function f
"
);
c!(
    recursive_results_07,
    "recursive integer function f() result(r)
r=1
end function f
"
);
c!(
    recursive_results_08,
    "recursive real function f() result(r)
r=1.0
end function f
"
);
c!(
    recursive_results_09,
    "recursive integer function f(n) result(r)
integer::n
if (n==0) then
 r=0
else
 r=f(n-1)
end if
end function f
"
);
c!(
    recursive_results_10,
    "recursive integer function f(n) result(r)
integer::n
r=1
end function f
"
);
