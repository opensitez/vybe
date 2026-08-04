! vybe-test: fortran/intent_optional_extended/optional_subroutine_padding_present
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
call emit_padded(7, 5)
contains
subroutine emit_padded(v, pad)
integer, intent(in) :: v
integer, intent(in), optional :: pad
integer :: out
out = v
if (present(pad)) out = out + pad
if ((out) /= 12) then
    print *, "FAIL: want [12] got [", out, "]"
    stop 1
end if
end subroutine emit_padded
end program t
