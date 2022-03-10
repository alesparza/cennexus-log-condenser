# cennexus-log-condenser

A small project to clean up Cennexus Host logs.

## Dependencies

* [Python 3.10.2](https://www.python.org/)
* [tqdm 4.63.0](https://github.com/tqdm/tqdm)
* [openpyxl 3.0.9](https://pypi.org/project/openpyxl/)

## Installation

Be sure to install the dependencies

```
pip install tqdm
pip install openpyxl
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

I don't have this set up yet.  So I guess default copywrite rules apply.

## Contributors

alesparza[https://github.com/alesparza)


