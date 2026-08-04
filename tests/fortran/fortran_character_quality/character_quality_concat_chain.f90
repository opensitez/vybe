! vybe-test: fortran/fortran_character_quality/character_quality_concat_chain
! origin: languages/fortran/tests/fortran/test_fortran_character_quality.rs

program character_quality_concat_chain
    character(len=30) :: text
    text = 'foo' // '-' // 'bar' // '-' // 'baz'
    if (trim(trim(text)) /= "foo-bar-baz") then
    print *, "FAIL: want [foo-bar-baz] got [", trim(text), "]"
    stop 1
end if
    if ((len_trim(text)) /= 11) then
    print *, "FAIL: want [11] got [", len_trim(text), "]"
    stop 1
end if
end program character_quality_concat_chain
