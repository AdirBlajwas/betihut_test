# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import betihut, sys





def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


def main(argv):
    print(argv[0])
    df = betihut.load_data(argv[1])
    df = betihut.clean_data(df)
    betihut.forms(df)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    main(sys.argv)
