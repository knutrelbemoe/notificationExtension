import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';
import * as React from 'react';
import './notify.css';
import ShowItemDetails from './ShowItemDetails';



interface WelcomeProps {
    name: string,
    items: any
    ctx: ApplicationCustomizerContext
}
export const NotificationDetails: React.SFC<WelcomeProps> = (props) => {
    var divItems = [];

    if (props.items.length > 0) {
        props.items.forEach(element => {
            divItems.push(
                <div className="flexWarperitem MNcard MNcard-link" onClick={(e) => ShowDetail(e)}>
                    <h4 style={{ margin: '0' }}>{element.Title}</h4>
                    <p>Start: {GetDate(element.StartDate)}</p>
                    <p>Slutt: {GetDate(element.EndDate)}</p>

                    <div className="hidden" style={{ display: 'none' }}>
                        <span>{element.Title}</span>
                        <span>{GetDate(element.StartDate)}</span>
                        <span>{GetDate(element.EndDate)}</span>
                        <span>{element.Description}</span>
                        <span>{element.SalesNo}</span>
                    </div>

                    <div className="icon">
                        <div className="arrow"></div>
                    </div>
                </div>
            );
        });
    }

    return (
        <div className="flexWarper">
            {divItems}
        </div>
    );
}

function ShowDetail(e) {
    var popupElements;

    if (e.target.classList.length > 0) {
        if (e.target.classList[0] == "flexWarperitem") {
            console.log('1');
            popupElements = e.target.getElementsByClassName("hidden")[0].childNodes;
        }
        else {
            console.log('2');
            popupElements = e.target.parentElement.closest("div.flexWarperitem").getElementsByClassName("hidden")[0].childNodes;
        }
    }
    else {
        console.log('3');
        popupElements = e.target.parentElement.closest("div.flexWarperitem").getElementsByClassName("hidden")[0].childNodes;
    }

    console.log(popupElements);

    const dialog: ShowItemDetails = new ShowItemDetails();
    dialog.message = 'Show details:';

    dialog.htmlElements = popupElements;
    dialog.show().then(() => {
        //
    });
}

function GetDate(strDt) {
    var formatDate = "";

  /*  if (strDt != null && strDt !== undefined) {
        var dt = new Date(strDt);
        var fMonth, fDay;
        var dd = dt.getDate();
        var mm = dt.getMonth() + 1;
        var yyyy = dt.getFullYear();

        if (dd < 10) {
            fDay = '0' + dd;
        }
        else {
            fDay = dd;
        }

        if (mm < 10) {
            fMonth = '0' + mm;
        }
        else {
            fMonth = mm;
        }
        var hrs = dt.getHours()+12;  
        formatDate = fDay + '.' + fMonth + '.' + yyyy + ' kl ' +
        hrs + ':' + dt.getMinutes();
    }*/
    var arr = strDt.split("T");
    var arrDate = arr[0].split("-");
    var arrTime = arr[1].split(":");
    const nuevo = arrTime.map((i) => Number(i));
    var hrs = nuevo[0]+2;
    var retDate = arrDate[2]+"."+arrDate[1]+"."+arrDate[0] + " kl " + hrs+":"+arrTime[1];
    return retDate;
}