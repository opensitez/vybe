! vybe-test: fortran/keyword_arguments/kw_25
! origin: languages/fortran/tests/fortran/test_keyword_arguments.rs
subroutine s(x, y)
real, intent(in) :: x
integer, intent(out) :: y
end
program p
integer :: y
call s(3.14, y)
end program p
