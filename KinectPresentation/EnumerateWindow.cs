using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KinectPresentation
{
    class EnumerateWindow
    {
        public IList childWindows;
        public IntPtr handle;
        public StringBuilder caption;
        public StringBuilder className;

        public EnumerateWindow()
        {
            this.handle = (IntPtr)0;
            this.caption = null;
            this.className = null;
        }

        public EnumerateWindow(int _handle, String _caption, String _className)
        {
            this.handle = (IntPtr)_handle;
            this.caption = new StringBuilder(_caption);
            this.className = new StringBuilder(_className);
        }

        public EnumerateWindow AddChild(int _handle, String _caption, String _className)
        {
            EnumerateWindow child = new EnumerateWindow(_handle, _caption, _className);

            if (this.childWindows == null)
            {
                this.childWindows = new ArrayList();
            }
            this.childWindows.Add(child);

            return child;
        }

        public Boolean hasChild()
        {
            if (this.childWindows == null)
            {
                return false;
            }
            else
            {
                return true;
            }
        }
    }
}
