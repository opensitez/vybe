! vybe-test: fortran/internal_io_extended/iio_write_to_character_array_slot
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=6) :: slots(3)
write(slots(2), '(I0)') 24
print *, trim(slots(2))
end program t
