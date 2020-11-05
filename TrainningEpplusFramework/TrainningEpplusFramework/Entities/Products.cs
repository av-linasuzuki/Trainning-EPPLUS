using System;
using System.Collections.Generic;
using System.Text;

namespace TrainningEpplusFramework.Entities
{
    class Products
    {
        public int productId{ get; set; }
        public string productName { get; set; }
        public double productPriceUnit { get; set; }
        public int productAmount { get; set; }
        public double productSubtotal { get; set; }

        public Products()
        {

        }
        public Products(int prodId, string prodName, double prodPriceUnit, int prodAmount, double prodSubtotal)
        {
            this.productId = prodId;
            this.productName = prodName;
            this.productPriceUnit = prodPriceUnit;
            this.productAmount = prodAmount;
            this.productSubtotal = prodSubtotal;
        }
       
    }
}
