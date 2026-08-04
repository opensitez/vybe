! vybe-test: fortran/oop/oop_class_arg_07
! origin: languages/fortran/tests/fortran/test_oop.rs
subroutine s(x)
class(*), intent(in) :: x
end subroutine s
