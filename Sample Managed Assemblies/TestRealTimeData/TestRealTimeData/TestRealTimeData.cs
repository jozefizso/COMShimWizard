using System;
using System.Timers;
using System.Runtime.InteropServices;
using System.Collections;

using Excel = Microsoft.Office.Interop.Excel;


namespace TestRealTimeData
{
	[ComVisible(true)]
	[GuidAttribute("4B418B7C-0547-4f99-A102-0EDC3E84E0C9")]
	[ProgId("TestRealTimeData.SomeProgId")]
	public class Stock : Excel.IRtdServer
	{

		// This will hold our collection of stock items.
		private ArrayList dataCollection;


		// Default constructor required for COM object.
		public Stock()
		{
			dataCollection = new ArrayList();
		}


		// IRtdServer methods. ////////////////////////////////////////////////

		// Return 1 to tell Excel we're still alive.
		public int Heartbeat()
		{
			return 1;
		}

		// This will be called just after the constructor. We need to cache the
		// IRTDUpdateEvent reference. We'll also setup a timer (with a 1000
		// millisecond interval) to trigger the simulated data updates. Finally, 
		// return 1 if all is OK.
		public int ServerStart(
			Microsoft.Office.Interop.Excel.IRTDUpdateEvent CallbackObject)
		{
			xlUpdateEvent = CallbackObject;

			timer = new Timer(1000);
			timer.AutoReset = true;
			timer.Elapsed += new ElapsedEventHandler(TimerEventHandler);

			return 1;
		}

		// Clean up, by setting the cached IRTDUpdateEvent reference to null,
		// stopping the timer, and setting the reference to null.
		public void ServerTerminate()
		{
			xlUpdateEvent = null;
			if (timer.Enabled)
			{
				timer.Stop();
			}
			timer = null;
		}

		// This is called when a file is opened that contains real-time data 
		// functions or when a user types in a new formula which contains 
		// the RTD function.
		public object ConnectData(
			int TopicID, ref System.Array Strings, ref bool GetNewValues)
		{
			// Make sure the timer has been started.
			if (!timer.Enabled)
			{
				timer.Start();
			}

			// Set GetNewValues to true to indicate that new values 
			// will be acquired.
			GetNewValues = true;

			// The array of strings passed in will be the parameters
			// to the RTD function in the worksheet. In our example, we're
			// only expecting one string - and this will be the stock name.
			string stockName = (string)Strings.GetValue(0);

			// Check to see if the requested topic is already in our
			// collection. If not, create a new data item to represent
			// it, set its TopicID from the valued passed in by Excel,
			// and add the item to our collection.
			StockData dataItem = null;
			if (!dataCollection.Contains(stockName))
			{
				dataItem = new StockData(stockName);
				dataItem.TopicID = TopicID;
				dataCollection.Add(dataItem);
			}
			else
			{
				foreach (StockData sd in dataCollection)
				{
					if (sd.Name == stockName)
					{
						dataItem = sd;
						break;
					}
				}
			}

			return dataItem.Price;
		}

		// Called from Excel for each previously connected use of the RTD 
		// function in the worksheet. We'll implement this to walk the 
		// collection of data items, and if we find one that matches the
		// specified TopicID, we'll remove it from the collection.
		public void DisconnectData(int TopicID)
		{
			foreach (StockData dataItem in dataCollection)
			{
				if (dataItem.TopicID == TopicID)
				{
					dataCollection.Remove(dataItem.Name);
				}
			}

			// If we've emptied the collection, stop the timer.
			if (dataCollection.Count == 0 && timer.Enabled)
			{
				timer.Stop();
			}
		}

		// Excel will call this method when it is ready to receive our data.
		// We must change the value of the TopicCount to the number of elements
		// in the array that we return. The data returned to Excel will be
		// a Variant containing a two-dimensional array. The first dimension
		// represents the list of topic IDs. The second dimension represents 
		// the values associated with the topic IDs.
		public System.Array RefreshData(ref int TopicCount)
		{
			// Declare a 2D array for the return value.
			object[,] variants = new object[2, dataCollection.Count];

			int itemCount = 0;
			for ( ; itemCount < dataCollection.Count; itemCount++)
			{
				StockData dataItem = (StockData)dataCollection[itemCount];
				variants[0, itemCount] = dataItem.TopicID;
				variants[1, itemCount] = dataItem.Price;
			}

			TopicCount = itemCount+1;
			return variants;
		}

		// ////////////////////////////////////////////////////////////////////

		// Excel exposes the IRTDUpdateEvent interface, so that we can
		// call into its UpdateNotify method to tell Excel that we have
		// new data ready.
		private Excel.IRTDUpdateEvent xlUpdateEvent;


		// Timer to trigger updating the data.
		private Timer timer;

		// Handler for the timer events - we use this to update the data.
		private void TimerEventHandler(object sender, ElapsedEventArgs e)
		{
			// Update the data.
			Random random = new Random();
			foreach (StockData dataItem in dataCollection)
			{
				dataItem.Update(random);
			}

			// Tell Excel we have updated data available.
			xlUpdateEvent.UpdateNotify();
		}

	}

}



























