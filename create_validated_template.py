"""Create a blank validated MRRO publisher-submission workbook."""

from argparse import ArgumentParser

from validated_template import create_submission_workbook


def main() -> None:
    parser = ArgumentParser()
    parser.add_argument("output", help="Path for the .xlsx template")
    args = parser.parse_args()
    create_submission_workbook(args.output)


if __name__ == "__main__":
    main()
