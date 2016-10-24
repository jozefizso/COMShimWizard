using System;

namespace TestRealTimeData
{

	// We use this class to encapsulate the data.
	internal class StockData
	{
		// The stock name, eg "MSFT".
		private string name;
		public string Name
		{
			get { return name; }
		}

		// The stock current price.
		private double price;
		public double Price
		{
			get { return price; }
		}

		// A unique ID, which will be assigned by Excel.
		private int topicID = -1;
		public int TopicID
		{
			get { return topicID; }
			set { topicID = value; }
		}

		// In the constructor, we'll generate a random value to
		// use for the initial price.
		public StockData(string name)
		{
			this.name = name;
			Random random = new Random();
			price = random.NextDouble() * 100;
		}

		// This method simulates market movements in stock price,
		// using randomly-generated values. A second random value
		// is used to determine whether the simulated price change
		// should be an increase or decrease.
		public void Update(Random random)
		{
			double priceChange = random.NextDouble();
			if (random.Next(1, 10) < 5)
			{
				price -= priceChange;
			}
			else
			{
				price += priceChange;
			}
		}


		// Override Equals and GetHashCode so that we can
		// add instances of this type to a collection, and perform
		// operations such as Contains.
		public override bool Equals(object obj)
		{
			StockData tmp = (StockData)obj;
			if (tmp != null)
			{
				if (tmp.name == this.name)
				{
					return true;
				}
			}
			return false;
		}

		public override int GetHashCode()
		{
			return this.name.GetHashCode();
		}
	}
}






























