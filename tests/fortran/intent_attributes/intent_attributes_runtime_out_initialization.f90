! vybe-test: fortran/intent_attributes/intent_attributes_runtime_out_initialization
! origin: languages/fortran/tests/fortran/test_intent_attributes.rs

program test_intent_attributes
integer :: x
integer :: y = 2
call write_out(y, x)
if ((x) /= 4) then
    print *, "FAIL: want [4] got [", x, "]"
    stop 1
end if

contains
subroutine write_out(src, dst)
integer, intent(in) :: src
integer, intent(out) :: dst
dst = src * 2
end subroutine write_out
end program test_intent_attributes
