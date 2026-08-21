! vybe-test: fortran/flush_statement/compile_flush_output_unit_no_arg
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs

program t
    use iso_fortran_env, only: output_unit
    write(output_unit, *) 'line'
    flush(output_unit)
end program t
