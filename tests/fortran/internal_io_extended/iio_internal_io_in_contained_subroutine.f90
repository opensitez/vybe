! vybe-test: fortran/internal_io_extended/iio_internal_io_in_contained_subroutine
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
call emit(9)
contains
subroutine emit(n)
integer, intent(in) :: n
character(len=6) :: buf
write(buf, '(I0)') n
print *, trim(buf)
end subroutine emit
end program t
