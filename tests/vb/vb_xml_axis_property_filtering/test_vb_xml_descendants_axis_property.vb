' vybe-test: vb/vb_xml_axis_property_filtering/test_vb_xml_descendants_axis_property
' origin: languages/vb/tests/vb/test_vb_xml_axis_property_filtering.rs

Imports System.Xml.Linq

Module Program
    Sub Main()
        Dim catalog = <catalog>
                          <book id="1">
                              <title>Book One</title>
                          </book>
                          <book id="2">
                              <title>Book Two</title>
                          </book>
                      </catalog>

        Dim titles = catalog...<title>
        For Each t In titles
            Console.WriteLine(t.Value)
        Next
    End Sub
End Module
