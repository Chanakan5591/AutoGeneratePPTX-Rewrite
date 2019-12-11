from pptx import Presentation
import argparse
import sys

parser = argparse.ArgumentParser(description="This tool will help you create your presentation slide fast and easy. Created by Chanakan Mungtin")
parser.add_argument("-i","--input",help="Specify powerpoint draft file created by format within 'format.txt'", nargs='?', const=None, required=True)
parser.add_argument("-p","--pages",help="Specify how many page will be in the presentation -- DEFAULT 1", nargs='?', const=0, required=False, type=int)

args = parser.parse_args()

def inputFileRead(filename):
    try:
        f = open(filename, "r")
        fcontent = f.readline()
        return fcontent, f
    except IOError:
        print("Error! File inaccessible")
        sys.exit(1)

def main():
    if args.input != None:
        signCount = 0
        callCount = 0
        prs = Presentation()
        title_slide_layout = prs.slide_layouts[0]
        content_slide_layout = prs.slide_layouts[1]
        title_slide = prs.slides.add_slide(title_slide_layout)
        content_slide = []
        nn = 2
        for i in range(int(args.pages) - 1):
            content_slide.append(prs.slides.add_slide(content_slide_layout))
            n = int(len(content_slide))
            if nn == int(args.pages):
                count = list(range(n))
                nn = nn + 1
            else:
                nn = nn + 1
        title = title_slide.shapes.title
        subtitle = title_slide.placeholders[1]
        draft, f = inputFileRead(str(args.input))
        while draft:
            fileContent = str(draft.strip())
            if "###" in fileContent:
                title.text = str(fileContent.replace("###", ""))
                prs.save('output.pptx')
            if "***" in fileContent:
                subtitle.text = str(fileContent.replace("***", ""))
                prs.save('output.pptx')
            if "--------" in fileContent:
                callCount = callCount + 1
                if args.pages <= 1:
                    pass
                else:
                    if callCount > 1:
                        PCDraft = signCount + 1
                        signCount = signCount + 1
                    else:
                        PCDraft = 0

            if "%%%" in fileContent:
                try:
                    content_title = content_slide[PCDraft].shapes.title
                    content_title.text = str(fileContent.replace("%%%", ""))
                except(IndexError):
                    print("Pages is not equal to section in draft file. exiting...")
                    sys.exit(1)
                prs.save('output.pptx')

            if "p>" in fileContent:
                try:
                    content_con = content_slide[PCDraft].shapes.placeholders[1]
                    content_con.text = str(fileContent.replace("p>", ""))
                except(IndexError):
                    print("Pages is not enough. exiting...")
                    sys.exit(1)
                prs.save('output.pptx')


            draft = f.readline()
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
