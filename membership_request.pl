#!C:\strawberry\perl\bin\perl.exe
use strict;
use Win32::OLE;
my $Notes = Win32::OLE->new('Notes.NotesSession')
    or die "Cannot start Lotus Notes Session object.\n";
my ($Version) = ($Notes->{NotesVersion} =~ /\s*(.*\S)\s*$/);

#my $Database = $Notes->GetDatabase('', 'C:\Users\your_username\AppData\Local\Lotus\Notes\Data\mail\yourLotusNotes.nsf'); 
my $Database = $Notes->GetDatabase('', 'C:\Users\your_username\AppData\Local\Lotus\Notes\Data\mail\yourLotusNotes.nsf');
 
my $AllDocuments = $Database->AllDocuments;

my $viewname = "New Mingle Users";		## Look for folder “Mingle New Users” on LotusNotes
my $view = $Database->GetView($viewname);
my $doc = $view->GetFirstDocument;

#Date#				## grab today’s Date
my $Date=`date/T`;
chomp($Date);
$Date=~s/\r|\n//g;

my $todayDate;
my($month,$day,$year);
if($Date=~/(\w+)\s(\d+)\/(\d+)\/(\d+)/){
	$month=$2;$day=$3;$year=$4;
	$todayDate="$month\/$day\/$year";
}

#my $file="mingle_membership_request_"."$month"."_"."$year".".csv";  ## create log file “mingle_membership_request_$month_$year.csv
my $file="C:\\Users\\Mingle\\Google Drive\\mingle\\mingle_membership_request_"."$month"."_"."$year".".csv";
my %db_membership;					## store log/.csv file into hash %db_membership	

if(-e $file){
	open(OUTPUT,">>$file");
	open my $info, $file;
	while(defined(my $line=<$info>)){
		chomp($line);
		if($line=~/MingleName,Program,MingleUser_id,Date_added,LotusEmail_id/){
			next;
		}
		my($name,$program,$user_id,$date,$email_id)=split(/,/,$line);
		#print "$name -> $program -> $user_id -> $date -> $email_id\n";
		$db_membership{$email_id}="$name,$program,$user_id,$date";
	}
}else{
	open(OUTPUT,">$file");
	print OUTPUT "MingleName,Program,MingleUser_id,Date_added,LotusEmail_id\n";
}

#my @db=keys(%db_membership);		## uncomment to print data store in hash
#foreach my $item(@db){
#	print "$item $db_membership{$item}\n";
#}

#my $todayDate="06/17/2014";			##### uncomment check all emails for a given date in LotusNotes
my $LotusDate = $Notes->CreateDateTime($todayDate);		
# Use $StartDate or $DateRange here					
my $Entries = $view->GetAllEntriesByKey($LotusDate);	##Get Document that is received todayDate
my $Entry = $Entries->GetFirstEntry();

while($Entry){
	my $doc = $Entry->Document();
	my $email_id = $doc->NoteID;		## LotusNotes Email _id
	chomp($email_id);
	$email_id=~s/\r|\n//g;
	
	#my $create_date = $doc->{Created};
	#my $date = $create_date->Date;
	
	my $check_from=$doc->{From}->[0];
	# search LotusNotes for particular email 
	my $from="Mingle <your.mingle.support\@email.com>";
	
	#Parsing Subject Header
	my $subject = $doc->GetFirstItem('Subject')->{Text};	 
	#if(($subject=~/(.*)\swants to join your project\s(.*)/)&&($subject !~/^Re:.*/)&&($subject !~/^Fw:.*/)&&($doc->{From}->[0] =~/$from/)){
	#if(($subject=~/(.*)\swants to join your project\s(.*)/)&&($subject !~/^Re:.*/)&&($subject !~/^Fw:.*/)&&($date=~/$todayDate/)){
	if(($subject=~/(.*)\swants to join your project\s(.*)/)&&($subject !~/^Re:.*/)&&($subject !~/^Fw:.*/)){
		print "From: $check_from\n";
		#print "--$from--\n";
		my $username=$1;			## store $1 to $username
		my $program_name=$2;		## store $2 to program_name
		if($check_from =~/$from/){
			print "======SUBJECT ($email_id)=======\n";
			print $doc->GetFirstItem('Subject')->{Text},"\n";	
			my $body=$doc->GetFirstItem('Body')->{Text};
			$body=~s/\r|\n//g;
			if($body=~/http:\/\/(.*)\/projects\/(.*)\/team\/.*user_id=(\d+)/){		#Look for Body Text
				print "======BODY======\n";
				#print $doc->GetFirstItem('Body')->{Text},"\n\n";
				my $servername=$1;					# store $1 to $servername
				my $program=$2;					# store $2 to $program
				my $userId=$3;					# store $3 to $userId
				#print "server:$1\nprogram:$2\nuser_id:$3\n";

				#print "======Send Request======\n";
				if(!exists($db_membership{$email_id})){
					print "this is new Email_id: ($email_id)\n";
					print "$username,$program_name,$userId,$todayDate\n";
					if($servername=~/mingle.server_host/){ ## edit this line to target Mingle Server
							my $request="curl -X POST -d\"membership\[permission\]=full_member\""." http:\/\/mingle_user:mingle_password\@"."mingle"."\/projects\/"."$program"."\/team\/add_user_to_team\?user_id="."$userId";
						print "$request\n\n";    ## replace mingle_login:mingle_pwd
						my $send_request=`$request 2>&1`;	
					
						if($send_request=~/$username.* has been added to the .* team successfully/){
							print "$send_request\n";
							print OUTPUT "$username,$program_name,$userId,$todayDate,$email_id\n";	## if CURL is successful add to log else send error to Mingle Admin
						}else{
							$Database->OpenMail;
							my $Document = $Database->CreateDocument;

							$Document->{Form} = 'Memo';
							$Document->{SendTo} = ['your@email.com','your@email.com2'];  ## --> Send Request Error 
							$Document->{Subject} = 'Mingle Membership Request Error';
							my $Body = $Document->CreateRichTextItem("Body");
							$Body->AppendText("MingleName: $username\rProgram: $program_name\rMingleUserID: $userId\rDate: $todayDate\r\r$send_request");
							$Document->Send(0);
						}
					}		
				}elsif($db_membership{$email_id} eq "$username,$program_name,$userId,$todayDate"){;
					print "This email_id have been processed: ($email_id)\n\n";
				}				
			}			
		}
	}
	$Entry = $Entries->GetNextEntry($Entry);		# get next email_id
}
close(OUTPUT);
