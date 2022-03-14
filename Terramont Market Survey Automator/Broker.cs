using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
namespace Terramont_Market_Survey_Automator
{
	public class Broker
	{
		private string name;
		private string title;
		private string officePhone;
		private string cellPhone;
		private string email;
		private string profilePicture;
		
		public Broker()
		{
			
		}

		public string Name
		{
			get { return name; }
			set { name = value; }
		}

		public string Title
		{
			get { return title; }
			set { title = value; }
		}
		
		public string OfficePhone
		{
			get { return officePhone; }
			set { officePhone = value; }
		}

		public string CellPhone
		{
			get { return cellPhone; }
			set { cellPhone = value; }
		}

		public string Email
		{
			get { return email; }
			set { email = value; }
		}

		public string ProfilePicture
		{
			get { return profilePicture; }
			set { profilePicture = value; }
		}
	}
}
