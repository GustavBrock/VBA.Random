# VBA.Random

![Help](https://raw.githubusercontent.com/GustavBrock/VBA.Random/master/images/EE%20Header.png)

### Truly random numbers in VBA
The access to true randomness is important in certain areas of math and statistics. Previously, truly random numbers have been either difficult or expensive to retrieve, but with the API from [ETH Zürich](http://qrng.ethz.ch/http_api/), everyone can now retrieve a set of random numbers at will for free.

This service is a web API for the quantum random number generator [Quantis](http://www.idquantique.com/random-number-generation/) developed by the Swiss company [ID Qantique](http://www.idquantique.com/):

> *The Quantis device produced by ID Qantique makes use of the uncertainty of photons based on a Polarising Beam Splitter (PBS), which reflects vertically polarised photons and transmits horizontally polarised photons as illustrated in the figure:*

<center>![Polarising Beam Splitter](images/polarising_beam_splitter.png)</center>

Very innovative idea. Please study the links above for the details.

### Getting access

The first task is to get access to the API and pass information about which "kind" of random numbers we wish to retrieve and how many. The choice is between *integer* values and *decimal* values while the count can be from a single number to many thousands.

An account is not needed, and as the data to retrieve is in Json format, a compact function can be used. The only caveat is, that sometimes the service seems to time out, thus a loop to handle this is included in the function **RetrieveDataResponse**.

### Retrieve a batch of random numbers

Having the option to connect to the API, two functions are provided to retrieve batches of random numbers:

	QrnIntegers
	QrnDecimals	

They both return an array with the retrieved numbers, for example, to collect 30 numbers between 10 and 20:

	Dim RandomNumbers As Variant
	RandomNumbers = QrnIntegers(30, 10, 20)

### Retrieve one random number
Sometimes you just need a single number. However, it takes about the same time to retrieve one number as several hundreds, and it would be unfriendly to burden the free service with a large count of calls for single numbers.
Thus, functions for retrieval of single numbers are included that - behind the scene - retrieve a batch of numbers and return these one by one as your application will need them:

	QrnInteger
	QrnDecimal

### Substituting VBA.Rnd
Also, a direct replacement for VBA.Rnd is provided:

	RndQrn

It takes the same argument and values but - due to the higher resolution - returns a Double, not a Single.

### Code

Code has been tested with both 32-bit and 64-bit Microsoft Access 2019 and 365.

### Documentation

Full documentation can be found here:

![EE Logo](images/EE%20Logo.png) 

[Truly Random Numbers in VBA](https://www.experts-exchange.com/articles/34471/Truly-Random-Numbers-in-VBA.html?preview=kYXBu8KHTtA%3D)

Included is a Microsoft Access example application and Microsoft Excel example workbook.

<hr>

*If you wish to support my work or need extended support or advice, feel free to:*

<p>
[<img src="https://raw.githubusercontent.com/GustavBrock/VBA.Random/master/images/BuyMeACoffee.png">](https://www.buymeacoffee.com/gustav/)