# cennexus-log-condenser

A small project to clean up Cennexus Host logs.  Cennexus is a Beckman Coulter automation line product and sometimes I need to pull a ton of data and filter out everything that isn't an ASTM message.

## Dependencies

* [Python](https://www.python.org/)
* [tqdm](https://github.com/tqdm/tqdm)
* [openpyxl](https://pypi.org/project/openpyxl/)

## Installation

Be sure to install the dependencies

```
pip install -r requirements.txt
```

## Usage

Collect the appropriate Cennexus log (Debug and Host 1, 2, or 3).

Extract the contents into a new directory.

Run the script for a single file:

```
# runs the script for a specific file
python3 cennexus-log-condenser.py -i <inputfile> -o <outputfile>
```

Or process the entire directory:

```
# runs the script for an entire directory
python3 cennexus-log-condenser.py -d <directory>
```

## Contributing

Open a pull request I guess.

## License

GNU General Public License v3.0

## Contributors

[alesparza](https://github.com/alesparza)


