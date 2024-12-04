export class MarkingManager {
    constructor() { }

    public static GetMessageType(): Promise<Office.CoercionType> {
        var promise = new Promise<Office.CoercionType>((resolve, reject) => {
            Office.context.mailbox.item.body.getTypeAsync((asyncResult: Office.AsyncResult<Office.CoercionType>) => {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    reject(asyncResult.error);
                } else {
                    const messageType: Office.CoercionType = asyncResult.value;
                    resolve(messageType);
                }
            });
        });
        return promise;
    }

    public static GetMessageBody(coersionType: Office.CoercionType): Promise<string> {
        var promise = new Promise<string>((resolve, reject) => {
            Office.context.mailbox.item.body.getAsync(coersionType, {}, (asyncResult: Office.AsyncResult<string>) => {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    reject(asyncResult.error);
                } else {
                    const message: string = asyncResult.value;
                    resolve(message);
                }
            });
        });

        return promise;
    }

    public static GetMessageSubject(): Promise<string> {
        var promise = new Promise<string>((resolve, reject) => {
            Office.context.mailbox.item.subject.getAsync((asyncResult: Office.AsyncResult<string>) => {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    reject(asyncResult.error);
                } else {
                    const message: string = asyncResult.value;
                    resolve(message);
                }
            });
        });

        return promise;
    }

    public static SetMessageSubject(subject: string): Promise<boolean> {
        var promise = new Promise<boolean>((resolve, reject) => {
            Office.context.mailbox.item.subject.setAsync(
                subject,
                {},
                (asyncResult) => {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        reject(asyncResult.error);
                    } else {
                        resolve(true);
                    }
                }
            );
        });
        return promise;
    }

    public static SetMessageBody(body: string, coersionType: Office.CoercionType): Promise<boolean> {
        var promise = new Promise<true>((resolve, reject) => {
            Office.context.mailbox.item.body.setAsync(
                body,
                { coercionType: coersionType },
                (asyncResult) => {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        reject(asyncResult.error);
                    } else {
                        resolve(true);
                    }
                }
            );
        });
        return promise;
    }
}