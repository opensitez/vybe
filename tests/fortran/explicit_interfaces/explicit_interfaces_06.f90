! vybe-test: fortran/explicit_interfaces/explicit_interfaces_06
! origin: languages/fortran/tests/fortran/test_explicit_interfaces.rs
program p
interface
subroutine s(x)
integer::x
end subroutine s
end interface
call s(1)
end program p

subroutine s(x)
integer::x
end subroutine s
