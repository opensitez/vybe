! vybe-test: fortran/interfaces/if_recursive_result_11
! origin: languages/fortran/tests/fortran/test_interfaces.rs
recursive integer function f(n) result(r)
integer::n
r=1
end function f
