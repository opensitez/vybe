! vybe-test: fortran/if_blocks/single_line_if_print
! origin: languages/fortran/tests/fortran/test_if_blocks.rs
program t
if (1 > 0) print *, "inline"
end program t
