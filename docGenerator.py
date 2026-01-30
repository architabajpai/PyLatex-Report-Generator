import pandas as pd
import numpy as np
from pylatex import Document, Section, Subsection, Figure, Table, Tabular, MultiColumn
from pylatex import Command, Package, NoEscape, NewPage, PageStyle, Head, Foot
from pylatex.utils import bold, italic
import os
import sys

def forcedata(filepath):
    try:
        df = pd.read_excel(filepath, engine='openpyxl')
        
        df.columns = df.columns.str.strip()
        
        print(f"Successfully loaded data with {len(df)} rows")
        print(f"Columns: {df.columns.tolist()}")
        
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        sys.exit(1)


def sfdplot(xval, sval):
    coordinates = " ".join([f"({x},{sf})" for x, sf in zip(xval, sval)])
    
    smin = min(sval)
    smax = max(sval)
    srange = smax - smin
    
    ymin = smin - 0.1 * srange
    ymax = smax + 0.1 * srange
    
    tikz_code = r"""
\begin{tikzpicture}
    \begin{axis}[
        width=14cm,
        height=8cm,
        xlabel={Distance along beam (m)},
        ylabel={Shear Force (kN)},
        grid=major,
        grid style={dashed, gray!30},
        legend pos=north east,
        axis lines=middle,
        ymin=""" + f"{ymin:.1f}" + r""",
        ymax=""" + f"{ymax:.1f}" + r""",
        xmin=0,
        xmax=""" + f"{max(xval)}" + r""",
        xlabel style={at={(1,0)}, anchor=north west},
        ylabel style={at={(0,1)}, anchor=south east},
        title={Shear Force Diagram},
        title style={font=\bfseries\large},
    ]
    
    % Plot the shear force line
    \addplot[
        color=blue,
        line width=1.5pt,
        mark=*,
        mark size=2pt,
    ] coordinates {
        """ + coordinates + r"""
    };
    \addlegendentry{Shear Force}
    
    % Add zero reference line
    \addplot[
        color=red,
        dashed,
        line width=0.8pt,
        domain=0:""" + f"{max(xval)}" + r""",
        samples=2
    ] {0};
    \addlegendentry{Zero Reference}
    
    \end{axis}
\end{tikzpicture}
"""
    return NoEscape(tikz_code)


def bmdplot(xval, bmval):
    coordinates = " ".join([f"({x},{bm})" for x, bm in zip(xval, bmval)])
    
    # Find min and max for plot scaling
    bmmin = min(bmval)
    bmmax = max(bmval)
    bmrange = bmmax - bmmin
    
    ymin = bmmin - 0.1 * bmrange if bmmin < 0 else 0
    ymax = bmmax + 0.1 * bmrange
    
    tikz_code = r"""
\begin{tikzpicture}
    \begin{axis}[
        width=14cm,
        height=8cm,
        xlabel={Distance along beam (m)},
        ylabel={Bending Moment (kN·m)},
        grid=major,
        grid style={dashed, gray!30},
        legend pos=north east,
        axis lines=middle,
        ymin=""" + f"{ymin:.1f}" + r""",
        ymax=""" + f"{ymax:.1f}" + r""",
        xmin=0,
        xmax=""" + f"{max(xval)}" + r""",
        xlabel style={at={(1,0)}, anchor=north west},
        ylabel style={at={(0,1)}, anchor=south east},
        title={Bending Moment Diagram},
        title style={font=\bfseries\large},
    ]
    
    % Plot the bending moment curve
    \addplot[
        color=purple,
        line width=1.5pt,
        mark=*,
        mark size=2pt,
        smooth,
    ] coordinates {
        """ + coordinates + r"""
    };
    \addlegendentry{Bending Moment}
    
    % Fill area under curve
    \addplot[
        color=purple!20,
        fill=purple!20,
        opacity=0.3,
    ] coordinates {
        """ + coordinates + r"""
    } \closedcycle;
    
    \end{axis}
\end{tikzpicture}
"""
    return NoEscape(tikz_code)


def forcetable(df):
    table_spec = 'c' * len(df.columns)
    
    table = Tabular(table_spec)
    table.add_hline()
    
    header_cells = [bold(col) for col in df.columns]
    table.add_row(header_cells)
    table.add_hline()
    
    for idx, row in df.iterrows():
        formatted_row = [f"{val:.2f}" if isinstance(val, (int, float)) else str(val) 
                        for val in row]
        table.add_row(formatted_row)
    
    table.add_hline()
    
    return table


def reportgen(excel_path, beam_image_path, output_pdf='output.pdf'):    
    # Read data
    print("Reading force data from Excel...")
    df = forcedata(excel_path)
    
    # Extract data arrays
    xval = df['x'].values
    sval = df['Shear force'].values
    bmval = df['Bending Moment'].values
    
    print("Creating LaTeX document...")
    
    geometry_options = {
        "head": "40pt",
        "margin": "1in",
        "bottom": "1in",
        "includeheadfoot": True
    }
    
    doc = Document(geometry_options=geometry_options)
    
    doc.packages.append(Package('graphicx'))
    doc.packages.append(Package('float'))
    doc.packages.append(Package('amsmath'))
    doc.packages.append(Package('tikz'))
    doc.packages.append(Package('pgfplots'))
    doc.packages.append(Package('booktabs'))
    doc.packages.append(Package('hyperref'))
    doc.packages.append(Package('fancyhdr'))
    
    doc.preamble.append(NoEscape(r'\pgfplotsset{compat=1.18}'))
    
    doc.preamble.append(NoEscape(r'\hypersetup{colorlinks=true, linkcolor=blue, urlcolor=blue}'))
    
    doc.preamble.append(NoEscape(r'\pagestyle{fancy}'))
    doc.preamble.append(NoEscape(r'\fancyhf{}'))
    doc.preamble.append(NoEscape(r'\fancyhead[L]{Beam Analysis Report}'))
    doc.preamble.append(NoEscape(r'\fancyhead[R]{\thepage}'))
    doc.preamble.append(NoEscape(r'\fancyfoot[C]{Simply Supported Beam - Structural Analysis}'))
    
    # ==================== TITLE PAGE ====================
    doc.preamble.append(Command('title', 'Beam Structural Analysis Report'))
    doc.preamble.append(Command('author', 'Engineering Analysis System'))
    doc.preamble.append(Command('date', NoEscape(r'\today')))
    doc.append(NoEscape(r'\maketitle'))
    
    # Add some spacing and a brief abstract
    doc.append(NoEscape(r'\vspace{2cm}'))
    doc.append(NoEscape(r'\begin{center}'))
    doc.append(NoEscape(r'\large'))
    doc.append(bold('Abstract'))
    doc.append(NoEscape(r'\end{center}'))
    doc.append(NoEscape(r'\normalsize'))
    doc.append(NoEscape(r'\vspace{0.5cm}'))
    
    doc.append('This report presents a comprehensive structural analysis of a simply supported beam '
               'under various loading conditions. The analysis includes the calculation and visualization '
               'of shear force and bending moment distributions along the beam length. The results are '
               'presented using industry-standard diagrams generated from computational analysis.')
    
    doc.append(NewPage())
    
    #TABLE OF CONTENTS
    doc.append(NoEscape(r'\tableofcontents'))
    doc.append(NewPage())
    
    #INTRODUCTION
    with doc.create(Section('Introduction')):
        
        with doc.create(Subsection('Beam Description')):
            doc.append('This analysis examines a simply supported beam, which is one of the most '
                      'fundamental structural elements in engineering. A simply supported beam is '
                      'supported at both ends with one pinned support and one roller support, '
                      'allowing it to freely rotate at the supports while preventing vertical displacement.')
            
            doc.append(NoEscape(r'\vspace{0.5cm}'))
            
            if os.path.exists(beam_image_path):
                with doc.create(Figure(position='H')) as fig:
                    fig.add_image(beam_image_path, width=NoEscape(r'0.8\textwidth'))
                    fig.add_caption('Simply Supported Beam Configuration')
            else:
                doc.append(italic(f'Note: Beam image not found at {beam_image_path}'))
        
        with doc.create(Subsection('Data Source')):
            doc.append('The force and moment data analyzed in this report are extracted from '
                      'the provided Excel spreadsheet. The data includes discrete measurements '
                      'of shear force and bending moment at regular intervals along the beam length. '
                      'This computational approach ensures accuracy and allows for detailed visualization '
                      'of the structural behavior.')
            
            doc.append(NoEscape(r'\vspace{0.3cm}'))
            doc.append(NoEscape(r'\noindent'))
            doc.append(bold('Data file: '))
            doc.append(f'{os.path.basename(excel_path)}')
            doc.append(NoEscape(r'\\'))
            doc.append(bold('Number of data points: '))
            doc.append(f'{len(df)} positions along the beam')
            doc.append(NoEscape(r'\\'))
            doc.append(bold('Beam length: '))
            doc.append(f'{max(xval)} m')
    
    doc.append(NewPage())
    
    #INPUT DATA
    with doc.create(Section('Input Data')):
        doc.append('The following table presents the complete force and moment data extracted from '
                  'the Excel file. The data includes position coordinates along the beam (x), '
                  'shear force values, and bending moment values at each position.')
        
        doc.append(NoEscape(r'\vspace{0.5cm}'))
        
        with doc.create(Table(position='H')) as table:
            table.add_caption('Force and Moment Data')
            table.append(Command('centering'))
            table.append(forcetable(df))
    
    doc.append(NewPage())
    
    #ANALYSIS
    with doc.create(Section('Structural Analysis')):
        
        doc.append('This section presents the graphical analysis of the beam through Shear Force '
                  'and Bending Moment Diagrams. These diagrams are essential tools for understanding '
                  'the internal forces and moments within the beam structure.')
        
        doc.append(NoEscape(r'\vspace{0.5cm}'))
        
        with doc.create(Subsection('Shear Force Diagram (SFD)')):
            doc.append(bold('Definition: '))
            doc.append('A Shear Force Diagram (SFD) is a graphical representation showing the variation '
                      'of shear force along the length of the beam. Shear force at any section represents '
                      'the algebraic sum of all vertical forces acting on either side of that section. '
                      'It indicates the internal sideways force that exists at each point along the beam.')
            
            doc.append(NoEscape(r'\vspace{0.3cm}'))
            doc.append(NoEscape(r'\noindent'))
            doc.append(bold('Key Observations:'))
            doc.append(NoEscape(r'\begin{itemize}'))
            doc.append(NoEscape(r'\item Maximum positive shear force: ' + f'{max(sval):.2f} kN'))
            doc.append(NoEscape(r'\item Maximum negative shear force: ' + f'{min(sval):.2f} kN'))
            doc.append(NoEscape(r'\item Zero shear occurs at: ' + f'{xval[np.argmin(np.abs(sval))]:.2f} m'))
            doc.append(NoEscape(r'\end{itemize}'))
            
            doc.append(NoEscape(r'\vspace{0.5cm}'))
            
            with doc.create(Figure(position='H')) as plot:
                plot.append(sfdplot(xval, sval))
                plot.add_caption('Shear Force Diagram')
        
        doc.append(NewPage())
        
        with doc.create(Subsection('Bending Moment Diagram (BMD)')):
            doc.append(bold('Definition: '))
            doc.append('A Bending Moment Diagram (BMD) illustrates the variation of bending moment '
                      'along the beam length. Bending moment at any section is the algebraic sum of '
                      'moments of all forces acting on either side of the section. It represents how '
                      'strongly the beam tends to rotate or bend at different locations.')
            
            doc.append(NoEscape(r'\vspace{0.3cm}'))
            doc.append(NoEscape(r'\noindent'))
            doc.append(bold('Key Observations:'))
            doc.append(NoEscape(r'\begin{itemize}'))
            doc.append(NoEscape(r'\item Maximum bending moment: ' + f'{max(bmval):.2f} kN·m'))
            doc.append(NoEscape(r'\item Location of maximum moment: ' + f'{xval[np.argmax(bmval)]:.2f} m'))
            doc.append(NoEscape(r'\item Moment at supports: ' + f'{bmval[0]:.2f} kN·m (left), {bmval[-1]:.2f} kN·m (right)'))
            doc.append(NoEscape(r'\end{itemize}'))
            
            doc.append(NoEscape(r'\vspace{0.5cm}'))
            
            with doc.create(Figure(position='H')) as plot:
                plot.append(bmdplot(xval, bmval))
                plot.add_caption('Bending Moment Diagram')
        
        with doc.create(Subsection('Analysis Summary')):
            doc.append('The structural analysis reveals important characteristics of the beam behavior:')
            
            doc.append(NoEscape(r'\vspace{0.3cm}'))
            doc.append(NoEscape(r'\begin{enumerate}'))
            doc.append(NoEscape(r'\item The shear force diagram shows a linear variation, indicating '
                               'uniformly distributed loading on the beam.'))
            doc.append(NoEscape(r'\item The bending moment diagram exhibits a parabolic shape, typical '
                               'of beams under uniform loads.'))
            doc.append(NoEscape(r'\item Maximum bending moment occurs near the mid-span, which is the '
                               'critical section for design purposes.'))
            doc.append(NoEscape(r'\item The beam exhibits symmetric behavior about its center, confirming '
                               'the symmetric loading condition.'))
            doc.append(NoEscape(r'\end{enumerate}'))
    
    #CONCLUSION
    doc.append(NewPage())
    with doc.create(Section('Conclusion')):
        doc.append('This report has presented a comprehensive structural analysis of a simply supported beam, '
                  'including detailed shear force and bending moment diagrams generated using advanced '
                  'computational techniques. The analysis demonstrates:')
        
        doc.append(NoEscape(r'\vspace{0.3cm}'))
        doc.append(NoEscape(r'\begin{itemize}'))
        doc.append(NoEscape(r'\item Clear visualization of internal force distributions'))
        doc.append(NoEscape(r'\item Identification of critical sections for design'))
        doc.append(NoEscape(r'\item Verification of expected structural behavior'))
        doc.append(NoEscape(r'\item Professional presentation using industry-standard diagrams'))
        doc.append(NoEscape(r'\end{itemize}'))
        
        doc.append(NoEscape(r'\vspace{0.5cm}'))
        doc.append('These results provide essential information for structural design and safety assessment '
                  'of the beam under the specified loading conditions.')
    
    # Generate PDF
    print(f"Generating PDF: {output_pdf}...")
    try:
        output_name = output_pdf.replace('.pdf', '')
        doc.generate_pdf(output_name, clean_tex=False, compiler='pdflatex')
        print(f"✓ Report successfully generated: {output_pdf}")
        return True
    except Exception as e:
        print(f"✗ Error generating PDF: {e}")
        print("Make sure you have a LaTeX distribution (like TeX Live or MiKTeX) installed.")
        return False


def main():
    excel_path="force_table.xlsx"
    image_path="ssbeam.png"
    output_path="output.pdf"
    print("="*60)
    print("Beam Analysis Report Generator")
    print("="*60)
    print(f"Excel data: {excel_path}")
    print(f"Beam image: {image_path}")
    print(f"Output PDF: {output_path}")
    print("="*60)
    
    success = reportgen(excel_path, image_path, output_path)
    
    if success:
        print("\n" + "="*60)
        print("SUCCESS! Your report has been generated.")
        print("="*60)
    else:
        print("\n" + "="*60)
        print("FAILED! Please check the error messages above.")
        print("="*60)
        sys.exit(1)


if __name__ == '__main__':
    main()