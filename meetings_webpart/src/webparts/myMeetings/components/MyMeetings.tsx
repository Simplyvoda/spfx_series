/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import styles from './MyMeetings.module.scss'; 
import { IMyMeetingsProps } from './IMyMeetingsProps';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react'; // ui packages for spinner
import * as moment from 'moment'; // import moment package to format time read moment docs to understand usage
import { MSGraphClientV3 } from '@microsoft/sp-http'; // required to connect to Microsoft Graph

// type definitions
interface IMeetingRes {
  odata: string;
  value: MeetingObj[];
}

interface MeetingObj {
  id: string;
  subject: string;
  start: timeobj;
  end: timeobj;
}

interface timeobj {
  dateTime: string;
  timeZone: string;
}

const MyMeetings = (props: IMyMeetingsProps): React.ReactElement => {
  const [meeting, setMeeting] = React.useState<Array<MeetingObj>>([]); // object to store response data
  const [isLoadingMeeting, setIsLoadingMeeting] = React.useState(true);


  // This function utilizes MSGraphClientV3 to connect with Microsoft Graph.
  // We access the msGraphClientFactory through the web part's context.
  // this function gets the logged in users meeting from outlook calender
  const getUserMeetings = async (): Promise<void> => {
    try {
      await props.spcontext.msGraphClientFactory //make sure you make changes to IMeetingsProps.ts to include spcontext
        .getClient('3')
        .then((client: MSGraphClientV3): void => {
          client
            .api(`/me/calendar/calendarView?startDateTime=${moment().utc().format()}&endDateTime=${moment().endOf('day').utc().format()}`)
            .select('id, subject, start, end')
            .orderby('start/dateTime')
            .get((err: unknown, res: IMeetingRes) => {
              if (err) {
                console.log("MSGraphAPI Error")
                console.error(err);
                return;
              }
              setMeeting(res.value);
            });
        })
    } catch (err) {
      console.log("MSGraphAPI Error")
      console.error(err);
    } finally {
      setIsLoadingMeeting(false);
    }
  }

  // use effect hook to fetch meetings each time meeting state is updated or the page reloads 
  React.useEffect(() => {
    getUserMeetings().catch(err => console.error(err));
  }, [meeting])

  return (
    <div className={styles.teamsSection}>
      <h1>My Meetings</h1>
      <p>{moment().format('ll')}</p>
      {
        isLoadingMeeting ? (
          <div style={{ height: '6rem', display: 'flex', justifyContent: 'center', alignContent: 'center', width: '100%' }}>
            <Spinner size={SpinnerSize.large} label='Loading your meetings...' />
          </div>
        ) : meeting.length === 0 ? (
          <div className={styles.emptyStateContainer}>
            <img src='https://bluechiptec.sharepoint.com/sites/BCTIntranet/Site%20Assets/meeting-interview-svgrepo-com.svg' alt=""
              width='50px' height='50px'
            />
            <p>No upcoming meetings</p>
          </div>
        ) : (
          <div className={styles.eventContainer}>
            {
              meeting.slice(0, 3).map((each, index) => {
                return (
                  <div key={index} className={styles.singleEvent}>
                    <div className={styles.timeSection}>
                      <img src="https://bluechiptec.sharepoint.com/sites/BCTIntranet/Site%20Assets/logos_microsoft-teams.svg" alt="time" />
                    </div>
                    <div className={styles.leftSection}>
                      {
                        each.subject.length > 70 ?
                          <h3>{each.subject.slice(0, 70)}...</h3> :
                          <h3>{each.subject}</h3>
                      }
                      <p>{moment(each.start.dateTime).add('1', 'hour').format('LT')}</p>
                    </div>
                  </div>
                )
              })
            }
          </div>
        )
      }
    </div>
  )
}
export default MyMeetings;
