use HTML::TableExtract;
use TEXT::Table;
use Text::Trim; 
use Excel::Writer::XLSX;
use HTML::Parser;
use File::Copy; 
#use HTML::TreeBuilder;
#use LWP::Simple;
#use HTML::TableContentParser;
$dir='/Users/yashuvinay/Desktop/Perl/Github/Input/';
opendir DIR, $dir or die "cannot open dir $dir: $!";
my @file= grep { $_ ne '.' && $_ ne '..' && $_ ne '.DS_Store' } readdir DIR;
closedir DIR;

foreach(@file)
{
    $inputFile=$_;
    $doc=$dir.$inputFile;

    $outFile = substr($inputFile, 0, index($inputFile, '.')).'.xlsx';
    $outDir='/Users/yashuvinay/Desktop/Perl/Github/Output/';
    $outFile = $outDir.$outFile;

    my $workbook = Excel::Writer::XLSX->new($outFile);
    $worksheet = $workbook->add_worksheet("sheet1");
    $format = $workbook->add_format();

    #$text =~ /^(.*)$/;
    #$text =~ /,([\w\s]+?),/;
    my $headersTable =  [ 'Series of Member Payment Dependent Notes',
                          'Aggregate principal amount of Notes offered',
                          'Aggregate principal amount of Notes sold',
                          'Stated interest rate',
                          'Service Charge',
                          'Sale and Original Issue Date',
                          'Initial maturity',
                          'Final maturity',
                          'Amount of corresponding member loan funded by Lending Club' ];
    $colexcel = 0;
    $rowexcel = 0;

    foreach ($columnInexcel => $headersTable)
    {
        $worksheet->write( $rowexcel, $colexcel, trim($headersTable), $format);
    }

    callTable($headersTable);

    my $filename = 'report.txt';
    open(my $fh, '>', $filename) or die "Could not open file '$filename' $!";

    my $text;
    my $p = HTML::Parser->new(text_h => [ sub {$text .= shift},
                                          'dtext']);
    $p->parse_file($doc);
        print $fh $text;
    close $fh;

     $colexcel=9;
     $rowexcel=1;
     $worksheet->write( 0, $colexcel, "Credit Score Range" );
     my $file="report.txt";
     open my $info, $file or die "Could not open $file: $!";
     while( my $line = <$info>)
     {
         $tempLine = $line;
         chomp $line;

         if("$line" =~ /Credit Score Range$\:/)
         {
             $temp = <$info> for 1;
             $worksheet->write( $rowexcel, $colexcel, trim($temp) );
             $rowexcel++;
         }
     }
     close $info;

     $colexcel++;
     $rowexcel=1;
     $worksheet->write( 0, $colexcel, "Earliest Credit Line" );
     open FH,"$file";
     while( <FH>)
     {
         $tempLine = $line = $_;
         chomp $line;
         if("$line" =~ /Earliest Credit Line$\:/)
         {
             $temp = <FH>;
             $worksheet->write( $rowexcel, $colexcel, trim($temp) );
             $rowexcel++;
         }
     }
     close (FH);

     $colexcel++;
     $rowexcel=1;
     $worksheet->write( 0, $colexcel, "Open Credit Lines" );
     open FH,"$file";
     while( <FH>)
     {
         $tempLine = $line = $_;
         chomp $line;
         if("$line" =~ /Open Credit Lines$\:/)
         {
             $temp = <FH>;
             $worksheet->write( $rowexcel, $colexcel, trim($temp) );
             $rowexcel++;
         }
     }
     close (FH);

     $colexcel++;
     $rowexcel=1;
     $worksheet->write( 0, $colexcel, "Total Credit Lines" );
     open FH,"$file";
     while( <FH>)
     {
         $tempLine = $line = $_;
         chomp $line;
         if("$line" =~ /Total Credit Lines$\:/)
         {
             $temp = <FH>;
             $worksheet->write( $rowexcel, $colexcel, trim($temp) );
             $rowexcel++;
         }
     }
     close (FH);

     $colexcel++;
     $rowexcel=1;
     $worksheet->write( 0, $colexcel, "Revolving Credit Balance" );
     open FH,"$file";
     while( <FH>)
     {
         $tempLine = $line = $_;
         chomp $line;
         if("$line" =~ /Revolving Credit Balance$\:/)
         {
             $temp = <FH>;
             $worksheet->write( $rowexcel, $colexcel, trim($temp) );
             $rowexcel++;
         }
     }
     close (FH);

     $colexcel++;
     $rowexcel=1;
     $worksheet->write( 0, $colexcel, "Revolving Line Utilization" );
     open FH,"$file";
     while( <FH>)
     {
         $tempLine = $line = $_;
         chomp $line;
         if("$line" =~ /Revolving Line Utilization$\:/)
         {
             $temp = <FH>;
             $worksheet->write( $rowexcel, $colexcel, trim($temp) );
             $rowexcel++;
         }
     }
     close (FH);

     $colexcel++;
     $rowexcel=1;
     $worksheet->write( 0, $colexcel, "Inquiries in the Last 6 Months" );
     open FH,"$file";
     while( <FH>)
     {
         $tempLine = $line = $_;
         chomp $line;
         if("$line" =~ /Inquiries in the Last 6 Months$\:/)
         {
             $temp = <FH>;
             $worksheet->write( $rowexcel, $colexcel, trim($temp) );
             $rowexcel++;
         }
     }
     close (FH);

     $colexcel++;
     $rowexcel=1;
     $worksheet->write( 0, $colexcel, "Accounts Now Delinquent" );
     open FH,"$file";
     while( <FH>)
     {
         $tempLine = $line = $_;
         chomp $line;
         if("$line" =~ /Accounts Now Delinquent$\:/)
         {
             $temp = <FH>;
             $worksheet->write( $rowexcel, $colexcel, trim($temp) );
             $rowexcel++;
         }
     }
     close (FH);

     $colexcel++;
     $rowexcel=1;
     $worksheet->write( 0, $colexcel, "Delinquent Amount" );
     open FH,"$file";
     while( <FH>)
     {
         $tempLine = $line = $_;
         chomp $line;
         if("$line" =~ /Delinquent Amount$\:/)
         {
             $temp = <FH>;
             $worksheet->write( $rowexcel, $colexcel, trim($temp) );
             $rowexcel++;
         }
     }
     close (FH);

     $colexcel++;
     $rowexcel=1;
     $worksheet->write( 0, $colexcel, "Delinquencies Last 2 yrs" );
     open FH,"$file";
     while( <FH>)
     {
         $tempLine = $line = $_;
         chomp $line;
         if("$line" =~ /Delinquencies \(Last 2 yrs\)$\:/)
         {
             $temp = <FH>;
             $worksheet->write( $rowexcel, $colexcel, trim($temp) );
             $rowexcel++;
         }
     }
     close (FH);

     $colexcel++;
     $rowexcel=1;
     $worksheet->write( 0, $colexcel, "Months Since Last Delinquency" );
     open FH,"$file";
     while( <FH>)
     {
         $tempLine = $line = $_;
         chomp $line;
         if("$line" =~ /Months Since Last Delinquency$\:/)
         {
             $temp = <FH>;
             $worksheet->write( $rowexcel, $colexcel, trim($temp) );
             $rowexcel++;
         }
     }
     close (FH);

     $colexcel++;
     $rowexcel=1;
     $worksheet->write( 0, $colexcel, "Public Records On File" );
     open FH,"$file";
     while( <FH>)
     {
         $tempLine = $line = $_;
         chomp $line;
         if("$line" =~ /Public Records On File$\:/)
         {
             $temp = <FH>;
             $worksheet->write( $rowexcel, $colexcel, trim($temp) );
             $rowexcel++;
         }
     }
     close (FH);

     $colexcel++;
     $rowexcel=1;
     $worksheet->write( 0, $colexcel, "Months Since Last Record" );
     open FH,"$file";
     while( <FH>)
     {
         $tempLine = $line = $_;
         chomp $line;
         if("$line" =~ /Months Since Last Record$\:/)
         {
             $temp = <FH>;
             $worksheet->write( $rowexcel, $colexcel, trim($temp) );
             $rowexcel++;
         }
     }
     close (FH);

     $colexcel++;
     $rowexcel=1;
     $worksheet->write( 0, $colexcel, "Application Type" );
     open FH,"$file";
     while( <FH>)
     {
         $tempLine = $line = $_;
         chomp $line;
         if("$line" =~ /Application Type$\:/)
         {
             $temp = <FH>;
             $worksheet->write( $rowexcel, $colexcel, trim($temp) );
             $rowexcel++;
         }
     }
     close (FH);

     $colexcel++;
     $rowexcel=1;
     $worksheet->write( 0, $colexcel, "Home ownership" );
     open FH,"$file";
     while( <FH>)
     {
         $tempLine = $line = $_;
         chomp $line;
         if("$line" =~ /Home ownership$\:/)
         {
             $temp = <FH>;
             $worksheet->write( $rowexcel, $colexcel, trim($temp) );
             $rowexcel++;
         }
     }
     close (FH);

     $colexcel++;
     $rowexcel=1;
     $worksheet->write( 0, $colexcel, "Job title" ); open FH,"$file";
     while( <FH>)
     {
         $tempLine = $line = $_;
         chomp $line;
         if("$line" =~ /Job title$\:/)
         {
             $temp = <FH>;
             $worksheet->write( $rowexcel, $colexcel, trim($temp) );
             $rowexcel++;
         }
     }
     close (FH);

     $colexcel++;
     $rowexcel=1;
     $worksheet->write( 0, $colexcel, "Length of employment" ); open FH,"$file";
     while( <FH>)
     {
         $tempLine = $line = $_;
         chomp $line;
         if("$line" =~ /Length of employment$\:/)
         {
             $temp = <FH>;
             $worksheet->write( $rowexcel, $colexcel, trim($temp) );
             $rowexcel++;
         }
     }
     close (FH);

     $colexcel++;
     $rowexcel=1;
     $worksheet->write( 0, $colexcel, "Joint Debt to Income" ); open FH,"$file";
     while( <FH>)
     {
         $tempLine = $line = $_;
         chomp $line;
         if("$line" =~ /Joint Debt\-to\-Income$\:/)
         {
             $temp = <FH>;
             $worksheet->write( $rowexcel, $colexcel, trim($temp) );
             $rowexcel++;
         }
     }
     close (FH);

     $colexcel++;
     $rowexcel=1;
     $worksheet->write( 0, $colexcel, "Location" ); open FH,"$file";
     while( <FH>)
     {
         $tempLine = $line = $_;
         chomp $line;
         if("$line" =~ /Location$\:/)
         {
             $temp = <FH>;
             $worksheet->write( $rowexcel, $colexcel, trim($temp) );
             $rowexcel++;
         }
     }
     close (FH);

     $colexcel++;
     $rowexcel=1;
     $worksheet->write( 0, $colexcel, "Gross income" ); open FH,"$file";
     while( <FH>)
     {
         $tempLine = $line = $_;
         chomp $line;
         if("$line" =~ /Gross income$\:/)
         {
             $temp = <FH>;
             $worksheet->write( $rowexcel, $colexcel, trim($temp) );
             $rowexcel++;
         }
     }
     close (FH);

     $colexcel++;
     $rowexcel=1;
     $worksheet->write( 0, $colexcel, "Debt to income ratio" ); open FH,"$file";
     while( <FH>)
     {
         $tempLine = $line = $_;
         chomp $line;
         if("$line" =~ /Debt\-to\-income ratio$\:/)
         {
             $temp = <FH>;
             $worksheet->write( $rowexcel, $colexcel, trim($temp) );
             $rowexcel++;
         }
     }
     close (FH);

     $colexcel++;
     $rowexcel=1;
     $worksheet->write( 0, $colexcel, "Joint Gross Income" ); open FH,"$file";
     while( <FH>)
     {
         $tempLine = $line = $_;
         chomp $line;
         if("$line" =~ /Joint Gross Income$\:/)
         {
             $temp = <FH>;
             $worksheet->write( $rowexcel, $colexcel, trim($temp) );
             $rowexcel++;
         }
     }
     close (FH);

    $worksheet = $workbook->add_worksheet("sheet2");
    my $headersTable1 =  [ 'Series of Member Payment Dependent Notes'];
    my $headersTable =  [ 'Question', 'Answer' ];
    $colexcel1 = 0;
    $colexcel = 1;
    $rowexcel = 0;

    foreach ($columnInexcel => $headersTable1)
    {
        $worksheet->write( $rowexcel, $colexcel, trim($headersTable), $format);
    }
    foreach ($columnInexcel => $headersTable)
    {
        $worksheet->write( $rowexcel, $colexcel, trim($headersTable), $format);
    }
    my $table_extract1 = HTML::TableExtract->new(headers => $headersTable1);
    my $table_extract1 = HTML::TableExtract->new(headers => $headersTable);
    $table_extract1->parse_file($doc);
    $coexcel=0; $rowexcel=1;
    foreach $table ($table_extract1->tables)
    {
        for my $row ($table->rows)
        {
            $worksheet->write( $rowexcel, $coexcel, $row );
            $rowexcel++;
        }
    }


}

sub callTable
{
    my $table_extract = HTML::TableExtract->new(headers => @_);
    $table_extract->parse_file($doc);
    $coexcel=0; $rowexcel=1;
    foreach $table ($table_extract->tables)
    {
        for my $row ($table->rows)
        {
            $worksheet->write( $rowexcel, $coexcel, $row );
            $rowexcel++;
        }
    }
}
