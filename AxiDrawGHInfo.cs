using System;
using System.Drawing;
using Grasshopper.Kernel;

namespace AxiDrawGH
{
    public class AxiDrawGHInfo : GH_AssemblyInfo
    {
        public override string Name
        {
            get
            {
                return "AxiDrawGH";
            }
        }
        public override Bitmap Icon
        {
            get
            {
                //Return a 24x24 pixel bitmap to represent this GHA library.
                return null;
            }
        }
        public override string Description
        {
            get
            {
                //Return a short string describing the purpose of this GHA library.
                return "";
            }
        }
        public override Guid Id
        {
            get
            {
                return new Guid("db0ec317-b4af-48da-9abc-536f2d20ecdc");
            }
        }

        public override string AuthorName
        {
            get
            {
                //Return a string identifying you or your company.
                return "";
            }
        }
        public override string AuthorContact
        {
            get
            {
                //Return a string representing your preferred contact details.
                return "";
            }
        }
    }
}
