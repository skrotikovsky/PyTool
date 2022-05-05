import sys

sys.path.append(r"C:\Program Files\Autodesk\Navisworks Manage 2020")
import clr

clr.AddReferenceToFile('Interop.NavisworksIntegratedAPI17.dll')
clr.AddReferenceToFile('Interop.NavisworksAutomationAPI17.dll')
import NavisworksIntegratedAPI17
import NavisworksAutomationAPI17

if __name__ == '__main__':
    # Create a new Navisworks Document, this will launch Navisworks
    m_doc = NavisworksAutomationAPI17.DocumentClass()
    # Read the state from the new Navisworks Document
    m_state = m_doc.State()
    # Tell Navisworks to be not visible
    m_doc.Visible = False
    # Open file
    filename = "D:\test.nwd"
    m_doc.OpenFile(filename)

    # get clash statistics

    m_clash = None
    for d_plugin in m_state.Plugins():
        if d_plugin.ObjectName == "nwOpClashElement":
            m_clash = d_plugin
            break

    if m_clash is None:
        print
        "Clash Detective not found"
    else:
        # Run all stored clash tests
        m_clash.RunAllTests()

        # Output results to list
        result = {}
        for clashTest in m_clash.Tests():
            check_name = clashTest.name
            result[check_name] = []
            for clashResult in clashTest.results():
                result[check_name].append(clashResult.name)