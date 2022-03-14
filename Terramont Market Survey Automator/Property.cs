using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Terramont_Market_Survey_Automator
{
	public class Property
	{
		private string address;
		private string landlord;
		private string rentableArea;
		private string term;
		private string occupancy;
		private string incentives;
		private string netRent;
		private string operationCosts;
		private string taxes;
		private string energy;
		private string totalAdditionalRent;
		private string grossRent;
		private string parking;
		private string comments;
		public Property()
		{

		}
		public string Address
		{
			get { return address; }
			set { address = value; }
		}
		public string Landlord
		{
			get { return landlord; }
			set { landlord = value; }
		}
		public string RentableArea
		{
			get { return rentableArea; }
			set { rentableArea = value; }
		}

		public string Term
		{
			get { return term; }
			set { term = value; }
		}
		public string Occupancy
		{
			get { return occupancy; }
			set { occupancy = value; }
		}
		public string Incentives
		{
			get { return incentives; }
			set { incentives = value; }
		} 
		public string NetRent
		{
			get { return netRent; }
			set { netRent = value; }
		}
		public string OperationCosts
		{
			get { return operationCosts; }
			set { operationCosts = value; }
		}
		public string Taxes
		{
			get { return taxes; }
			set { taxes = value; }
		}
		public string Energy
		{
			get { return energy; }
			set { energy = value; }
		}
		public string TotalAdditionalRent
		{
			get { return totalAdditionalRent; }
			set { totalAdditionalRent = value; }
		}
		public string GrossRent
		{
			get { return grossRent; }
			set { grossRent = value; }
		}
		public string Parking
		{
			get { return parking; }
			set { parking = value; }
		}
		public string Comments
		{
			get { return comments; }
			set { comments = value; }
		}


	}
}
