using System.Windows.Forms;

namespace bbOffice.Common
{
    public interface IShowOnPage
    {
        Form GetForm();
        string GetName();
    }
}
